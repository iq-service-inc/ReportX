﻿using ReportX.Rep.Attributes;
using ReportX.Rep.Model;
using ReportX.Rep.View.S5View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class KBStatic
    {

        protected string[] oldcols;
        protected string[] newcols;

        public string[] cols;
        private ModelKBStatic kbs;
        private List<ModelTR> trs;
        public MemberInfo[] modeli;
        public KBStatic(Type model)
        {
            trs = new List<ModelTR>();
            kbs = new ModelKBStatic();
            kbs.style = new ViewStyleKBStatic();

            List<MemberInfo> list_cols = new List<MemberInfo>();
            modeli = model.GetMembers();
            foreach (var member in model.GetMembers())
            {
                Present attr = member.GetCustomAttribute<Present>();
                if (attr == null) continue;

                int MetadataToken = member.MetadataToken,
                    inserted_index = 0;

                // sory by MetadataToken (declaration)
                for (int i = 0; i < list_cols.Count; i++)
                {
                    inserted_index = i;
                    if (MetadataToken < list_cols[i].MetadataToken) break;
                    inserted_index = i + 1;
                }
                list_cols.Insert(inserted_index, member);
            }
            string[] str_cols = new string[list_cols.Count]; //取得標題數量
            for (int i = 0; i < list_cols.Count; i++)
                str_cols[i] = list_cols[i].GetCustomAttribute<Present>().getName();//取得標題名稱
            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            kbs.colNum = cols.Length;

        }
        // 傳入一個陣列 
        public void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            kbs.colNum = cols.Length;
        }


        public string formatDate(DateTime dateTime1, DateTime dateTime2)
        {
            throw new NotImplementedException();
        }

        public void setTitle(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) kbs.author = author;
            if (company != null) kbs.company = company;
            if (sheetName != null) kbs.sheetName = sheetName;
            if (dateTime != null) kbs.datetime = dateTime;
            if (dateRange != null) kbs.dateRange = dateRange;
        }

        public void setCustomStyle(string css)
        {
            kbs.style.setCustomCSS(css);
        }

        public ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = kbs.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }

        public ModelTR appendRow(params object[] data)
        {
            ModelTR tr = new ModelTR();
            tr.tds = new List<ModelTD>();
            foreach (object cell in data)
            {

                ModelTD td = new ModelTD();
                var value = cell.GetType().GetProperty("value");

                if (value == null)
                {
                    td.data = cell.ToString();
                }
                else
                {
                    var colspan = cell.GetType().GetProperty("colspan");
                    var rowspan = cell.GetType().GetProperty("rowspan");
                    var fontSize = cell.GetType().GetProperty("fontSize");
                    var align = cell.GetType().GetProperty("align");
                    var bold = cell.GetType().GetProperty("bold");
                    var style = cell.GetType().GetProperty("style");
                    var className = cell.GetType().GetProperty("className");
                    var sum_c = cell.GetType().GetProperty("sum_c");
                    var sum_w = cell.GetType().GetProperty("sum_w");


                    if (value != null) td.data = value.GetValue(cell, null);
                    if (colspan != null) td.colspan = (int)colspan.GetValue(cell, null);
                    if (rowspan != null) td.rowspan = (int)rowspan.GetValue(cell, null);
                    if (fontSize != null) td.fontSize = fontSize.GetValue(cell, null).ToString();
                    if (align != null) td.align = align.GetValue(cell, null).ToString();
                    if (bold != null) td.bold = true;
                    if (style != null) td.style = style.GetValue(cell, null).ToString();
                    if (className != null) td.className = className.GetValue(cell, null).ToString();

                }
                tr.tds.Add(td);
            }
            trs.Add(tr);
            return tr;
        }

        public void appendTable<T>(T[] data, string trStyle = null, string className = null)
        {

            foreach (T tuple in data)
            {
                ModelTD[] tds = new ModelTD[cols.Length];
                ModelTR tr = new ModelTR();
                tr.tds = new List<ModelTD>();
                tr.style = trStyle;
                tr.className = className;
                foreach (var prop in tuple.GetType().GetProperties())
                {
                    try
                    {
                        Present attr = prop.GetCustomAttribute<Present>();
                        if (attr == null) continue;
                        int colinx = Array.IndexOf(cols, attr.getName());
                        object value = prop.GetValue(tuple, null);
                        tds[colinx] = new ModelTD()
                        {
                            col = attr.getName(),
                            data = value
                        };
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                foreach (ModelTD td in tds)
                    tr.tds.Add(td);
                trs.Add(tr);
            }
        }

        public int getColCount()
        {
            return cols.Length;
        }

        public string getsheetName()
        {
            return kbs.sheetName;
        }

        public string render(int? width = null)
        {

            kbs.body = new ViewBodyKBStatic(trs, width);
            ViewKBStatic report = new ViewKBStatic(kbs);
            return report.render();
        }

    }

}
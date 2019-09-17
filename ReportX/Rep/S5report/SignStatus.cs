﻿using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View.S5View;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class SignStatus: AbsOpenOffice
    {
        private ModelSignStatus ss;
        protected override string[] oldcols { get; set; }
        protected override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
        public MemberInfo[] modeli;
        public SignStatus(Type model)
        {
            trs = new List<ModelTR>();
            ss = new ModelSignStatus();
            ss.style = new ViewStyleSignStatus();

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
            ss.colNum = cols.Length;

        }
        // 傳入一個陣列 
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            ss.colNum = cols.Length;
        }

        public override void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) ss.author = author;
            if (company != null) ss.company = company;
            if (sheetName != null) ss.sheetName = sheetName;
            if (dateTime != null) ss.datetime = dateTime;
            if (dateRange != null) ss.dateRange = dateRange;
        }

        public override void setCustomStyle(string css)
        {
            ss.style.setCustomCSS(css);
        }

        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = ss.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }

        public override string render(int? width = null)
        { 

            ss.body = new ViewBodySignStatus(trs, width);
            ViewSignStatus report = new ViewSignStatus(ss);
            return report.render();
        }

    }
}

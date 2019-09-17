﻿using ReportX.Rep.Attributes;
using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ReportX.Rep.Common
{
    public abstract class AbsOffice: IReportX
    {
        protected abstract string[] oldcols { get; set; }
        protected abstract string[] newcols { get; set; }

        public abstract string[] cols { get; set; }
        public abstract string[] cut { get; set; }
        protected abstract List<ModelTR> trs { get; }

        public virtual string render(int? width = null, string File_type = null)
        {
            throw new NotImplementedException();
        }
        public virtual string render(int? width = null)
        {
            throw new NotImplementedException();
        }
        public abstract void changecut(string[] cut);
        public abstract void setCustomStyle(string css);
        public abstract ModelTR appendFullRow(string data, string trStyle = null, string className = null);
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
        public void appendTable(DataTable data, string trStyle = null, string className = null)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                ModelTD[] tds = new ModelTD[cols.Length];
                ModelTR tr = new ModelTR();
                tr.tds = new List<ModelTD>();
                tr.style = trStyle;
                tr.className = className;

                foreach (var prop in data.Columns)
                {
                    try
                    {
                        var column = prop;
                        int colinx = Array.IndexOf(cols, column.ToString());
                        if (colinx == -1) continue;
                        var value = data.Rows[i][column.ToString()];
                        tds[colinx] = new ModelTD()
                        {
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
        public abstract void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null);

    }
}
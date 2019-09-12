﻿using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Odf
{
    public class Ods:AbsOpenOffice
    {
        private ModelOds ods;

        protected override string[] oldcols { get; set; }
        protected override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }

        public Ods(Type model)
        {
            trs = new List<ModelTR>();
            ods = new ModelOds();
            ods.style = new ViewStyleOds();

            List<MemberInfo> list_cols = new List<MemberInfo>();

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

            string[] str_cols = new string[list_cols.Count];

            for (int i = 0; i < list_cols.Count; i++)
                str_cols[i] = list_cols[i].GetCustomAttribute<Present>().getName();


            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            ods.colNum = cols.Length;

        }
        public Ods(DataTable data)
        {
            trs = new List<ModelTR>();
            ods = new ModelOds();
            ods.style = new ViewStyleOds();


            string[] str_cols = new string[data.Columns.Count];

            for (int i = 0; i < data.Columns.Count; i++)
                str_cols[i] = data.Columns[i].ToString();


            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            ods.colNum = cols.Length;

        }
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            ods.colNum = cols.Length;
        }
        public void setOds(string author = null, string company = null, string sheetName = null)
        {
            if (author != null) ods.author = author;
            if (company != null) ods.company = company;
            if (sheetName != null) ods.sheetName = sheetName;
        }
        public override void setCustomStyle(string css)
        {
            ods.style.setCustomCSS(css);
        }
        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = ods.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }
        public override string render(int? width = null)
        {
            ods.body = new ViewBodyOds(trs, width);
            ViewOds report = new ViewOds(ods);
            return report.render();
        }

    }
}

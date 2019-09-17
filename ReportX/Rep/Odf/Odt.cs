using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Odf
{
    public  class Odt:AbsOpenOffice
    {

        private ModelOdt odt;

        protected override string[] oldcols { get; set; }
        protected override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
        public Odt(Type model)
        {
            trs = new List<ModelTR>();
            odt = new ModelOdt();
            odt.style = new ViewStyleOdt();

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
            odt.colNum = cols.Length;

        }
        public Odt(DataTable data)
        {
            trs = new List<ModelTR>();
            odt = new ModelOdt();
            odt.style = new ViewStyleOdt();


            string[] str_cols = new string[data.Columns.Count];

            for (int i = 0; i < data.Columns.Count; i++)
                str_cols[i] = data.Columns[i].ToString();


            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            odt.colNum = cols.Length;

        }
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            odt.colNum = cols.Length;
        }
        public override  void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) odt.author = author;
            if (company != null) odt.company = company;
            if (sheetName != null) odt.sheetName = sheetName;
        }
        public override void setCustomStyle(string css)
        {
            odt.style.setCustomCSS(css);
        }
        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = odt.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }
        public override string render(int? width = null)
        {
            odt.body = new ViewBodyOdt(trs  , width);
            ViewOdt report = new ViewOdt(odt);
            return report.render();
        }
        
    }
}

using ReportX.Rep.Attributes;
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
    public class KBStatic : AbsOpenOffice
    {


        private ModelKBStatic kbs;
        protected override string[] oldcols { get; set; }
        protected override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
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
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            kbs.colNum = cols.Length;
        }
        // 傳入一個陣列 
        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
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
        public override void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) kbs.author = author;
            if (company != null) kbs.company = company;
            if (sheetName != null) kbs.sheetName = sheetName;
            if (dateTime != null) kbs.datetime = dateTime;
            if (dateRange != null) kbs.dateRange = dateRange;
        }
        public override void setCustomStyle(string css)
        {
            kbs.style.setCustomCSS(css);
        }
        public override string render(int? width = null)
        {

            kbs.body = new ViewBodyKBStatic(trs, width);
            ViewKBStatic report = new ViewKBStatic(kbs);
            return report.render();
        }

    }

}

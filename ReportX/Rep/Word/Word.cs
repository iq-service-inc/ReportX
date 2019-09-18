using ReportX.Rep.Attributes;
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

namespace ReportX.Rep.Word
{
    public class Word : AbsOffice
    {
        //存取器
        protected override string[] oldcols { get; set; }
        protected override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
        private ModelWord word;
        private int colspan;
        public Word(DataTable model)
        {
            trs = new List<ModelTR>();
            word = new ModelWord();
            word.style = new ViewStyle();


            string[] str_cols = new string[model.Columns.Count];

            for (int i = 0; i < model.Columns.Count; i++)
                str_cols[i] = model.Columns[i].ToString();


            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            word.colNum = cols.Length;


        }
        public Word(Type model)
        {            
                trs = new List<ModelTR>();
                word = new ModelWord();
                word.style = new ViewStyle();

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
                word.colNum = cols.Length;
            
        }
        
        // 傳入一個陣列 
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            word.colNum = cols.Length;
        }

        public override void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) word.author = author;
            if (company != null) word.company = company;
            if (sheetName != null) word.sheetName = sheetName;
        }

        public override void setCustomStyle(string css)
        {
            word.style.setCustomCSS(css);
        }

        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = word.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }

        public override string  render(int? width = null)
        {
            word.body = new ViewBody(trs, width);
            ViewWord report = new ViewWord(word);
            return report.render();
        }
    }
}

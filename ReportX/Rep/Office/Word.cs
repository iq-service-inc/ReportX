using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System.Collections.Generic;
using System.Linq;

namespace ReportX.Rep.Office
{
    public class Word : AbsOffice
    {
        //存取器
        public override string[] oldcols { get; set; }
        public override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
        private ModelWord word;
        public Word()
        {
            trs = new List<ModelTR>();
            word = new ModelWord();
            word.style = new ViewStyle();
        }

        /// <summary>
        /// 過濾顯示欄位，需要在 setData 之後才能呼叫
        /// </summary>
        /// <param name="cut">需要顯示的欄位陣列</param>
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

        public override string render(int? width = null)
        {
            word.body = new ViewBody(trs, width);
            ViewWord report = new ViewWord(word);
            return report.render();
        }

        protected override void setReportColNum()
        {
            word.colNum = cols.Length;
        }
    }
}

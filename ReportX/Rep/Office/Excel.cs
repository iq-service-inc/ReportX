using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System.Collections.Generic;
using System.Linq;

namespace ReportX.Rep.Office
{
    public class Excel: AbsOffice
    {
        //存取器
        public override string[] oldcols { get; set; }
        public override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }
        private ModelExcel excel;

        public Excel()
        {
            trs = new List<ModelTR>();
            excel = new ModelExcel();
            excel.style = new ViewStyle();
        }
   
        // 傳入一個陣列 
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            excel.colNum = cols.Length;
        }

        public override void  setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
        {
            if (author != null) excel.author = author;
            if (company != null) excel.company = company;
            if (sheetName != null) excel.sheetName = sheetName;
        }

        public override void setCustomStyle(string css)
        {
            excel.style.setCustomCSS(css);
        }

        public override ModelTR appendFullRow(string data, string trStyle = null, string className = null)
        {
            ModelTR tr = new ModelTR();
            ModelTD td = new ModelTD();
            tr.tds = new List<ModelTD>();
            td.data = data;
            td.className = className;
            td.style = trStyle;
            td.colspan = excel.colNum;
            tr.tds.Add(td);
            trs.Add(tr);
            return tr;
        }
        public string getsheetName()
        {
            return excel.sheetName;
        }      
        
        public override string render(int? width = null)
        {
            excel.body = new ViewBody(trs, width);
            ViewExcel report = new ViewExcel(excel);
            return report.render();
        }

        protected override void setReportColNum()
        {
            excel.colNum = cols.Length;
        }

    }
}

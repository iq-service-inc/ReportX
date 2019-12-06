using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System.Collections.Generic;
using System.Linq;

namespace ReportX.Rep.OpenOffice.Ods
{
    public class Ods:AbsOpenOffice
    {
        private ModelOds ods;

        public override string[] oldcols { get; set; }
        public override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }

        public Ods()
        {
            trs = new List<ModelTR>();
            ods = new ModelOds();
            ods.style = new ViewStyleOds();
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
            ods.colNum = cols.Length;
        }

        public override void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)
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

        protected override void setReportColNum()
        {
            ods.colNum = cols.Length;
        }

        /// <summary>
        ///  Ods file 專用 Meta 宣告，用於 META-INF 檔案建立時填入
        /// </summary>
        public override string meta => "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.spreadsheet'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";
    }
}

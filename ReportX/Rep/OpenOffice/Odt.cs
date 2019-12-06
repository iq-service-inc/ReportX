using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System.Collections.Generic;
using System.Linq;

namespace ReportX.Rep.OpenOffice.Odt
{
    public  class Odt:AbsOpenOffice
    {
        private ModelOdt odt;

        public override string[] oldcols { get; set; }
        public override string[] newcols { get; set; }
        protected override List<ModelTR> trs { get; }
        public override string[] cols { get; set; }


        public Odt()
        {
            trs = new List<ModelTR>();
            odt = new ModelOdt();
            odt.style = new ViewStyleOdt();
        }
      
        public override void changecut(string[] cut)
        {
            newcols = cut;
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
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
            odt.body = new ViewBodyOdt(trs, cols.Length);
            ViewOdt report = new ViewOdt(odt);
            return report.render();
        }

        protected override void setReportColNum()
        {
            odt.colNum = cols.Length;
        }

        /// <summary>
        ///  Odt file 專用 Meta 宣告，用於 META-INF 檔案建立時填入
        /// </summary>
        public override string meta => "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.text'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='settings.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";

    }
}

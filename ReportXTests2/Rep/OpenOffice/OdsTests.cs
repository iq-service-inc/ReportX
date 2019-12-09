using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Model;
using ReportXTests2;

namespace ReportX.Rep.OpenOffice.Tests
{
    [TestClass()]
    public class OdsTests
    {
        [TestMethod()]
        public void renderTest()
        {

            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();

            Ods report = new Ods();
            report.setCol(data);
            if (cols != null && cols.Length > 0) report.changecut(cols);

            report.appendFullRow("測試增加一個 Header", "TableCellData", "Title");
            report.appendFullRow("測試客製化 Word", "TableCellData", "TitleTimeWord");
            report.appendFullRow("測試 CSS", "TableCellData", "TitleDateWord");

            report.setCustomStyle(customOfficeCSS);

            // 測試寫入欄位
            ModelTR col = report.appendRow(report.cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";

            report.appendTable(data);

            string res = report.render();
            string fileName = sampleData.FileName + ".ods";
            ReportSaver.saveOpenOfficeReport(fileName, res, report.meta);

            Assert.IsNotNull(res);
        }

        const string customOfficeCSS = @"<office:automatic-styles>
            <style:style style:name='ColumnWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#555555' style:repeat-content='false'/>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties fo:color='#FFFFFF' style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
            </style:style>
            <style:style style:name='Word' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
            </style:style>
            <style:style style:name='TotalWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDDDDD'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
            </style:style>
            <style:style style:name='TitleWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold'/>
            </style:style>
            <style:style style:name='DateRangeWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='15pt' style:font-size-asian='15pt' style:font-size-complex='15pt' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold'/>
            </style:style>
            <style:style style:name='CreaterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
              <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties fo:color='#555555' style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='11pt' style:font-size-asian='11pt' style:font-size-complex='11pt'/>
            </style:style>
            <style:style style:name='TableColumn' style:family='table-column'>
              <style:table-column-properties />
            </style:style>
            <style:style style:name='TableRow' style:family='table-row'>
              <style:table-row-properties style:row-height='auto' style:use-optimal-row-height='false' fo:break-before='auto'/>
            </style:style>
            <style:style style:name='ta1' style:family='table' style:master-page-name='mp1'>
              <style:table-properties table:display='true' style:writing-mode='lr-tb'/>
            </style:style>
            <style:page-layout style:name='pm1'>
              <style:page-layout-properties fo:margin-top='0.5in' fo:margin-bottom='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' style:print-orientation='portrait' style:print-page-order='ttb' style:first-page-number='continue' style:scale-to='100%' style:table-centering='none' style:print='objects charts drawings'/>
              <style:header-style>
                <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-bottom='0in'/>
              </style:header-style>
              <style:footer-style>
                <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-top='0in'/>
              </style:footer-style>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name='mp1' style:page-layout-name='pm1'>
              <style:header/>
              <style:header-left style:display='false'/>
              <style:footer/>
              <style:footer-left style:display='false'/>
            </style:master-page>
          </office:master-styles>
            ";
    }
}
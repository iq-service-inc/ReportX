using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Model;
using ReportXTests2;

namespace ReportX.Rep.OpenOffice.Odt.Tests
{
    [TestClass()]
    public class OdtTests
    {
        [TestMethod()]
        public void renderTest()
        { 
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();

            Odt report = new Odt();
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
            string fileName = sampleData.FileName + ".odt";
            ReportSaver.saveOpenOfficeReport(fileName, res, report.meta);

            Assert.IsNotNull(res);
        }

        const string customOfficeCSS = @"<office:automatic-styles>
            <style:style style:name='TableColumn' style:family='table-column'>
              <style:table-column-properties style:column-width='auto'/>
            </style:style>
            <style:style style:name='Table' style:family='table' style:master-page-name='MP0'>
              <style:table-properties  fo:margin-left='0in' table:align='center'/>
            </style:style>
            <style:style style:name='TableRow' style:family='table-row'>
              <style:table-row-properties/>
            </style:style>
            <style:style style:name='TableCellData' style:family='table-cell'>
              <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#DDEEFF' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
            </style:style>
            <style:style style:name='Title' style:parent-style-name='內文' style:family='paragraph'>
              <style:paragraph-properties fo:widows='2' fo:orphans='2' fo:break-before='page' fo:text-align='center'/>
            </style:style>
            <style:style style:name='TitleWord' style:parent-style-name='預設段落字型' style:family='text'>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt'/>
            </style:style>
            <style:style style:name='TitleDateWord' style:parent-style-name='內文' style:family='paragraph'>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold' fo:font-size='15pt' style:font-size-asian='15pt' style:font-size-complex='15pt'/>
            </style:style>
            <style:style style:name='TitleTimeWord' style:parent-style-name='內文' style:family='paragraph'>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:color='#555555' fo:font-size='10.5pt' style:font-size-asian='10.5pt' style:font-size-complex='10.5pt'/>
            </style:style>
            <style:style style:name='TitleCell' style:family='table-cell'>
              <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#555555' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
            </style:style>
            <style:style style:name='TitleCellWord' style:parent-style-name='內文' style:family='paragraph'>
              <style:paragraph-properties fo:text-align='center'/>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:color='#FFFFFF'/>
            </style:style>
            <style:style style:name='CellWord' style:family='table-cell'>
              <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
            </style:style>
            <style:style style:name='Word' style:parent-style-name='內文' style:family='paragraph'>
              <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體'/>
            </style:style>
            <style:style style:name='TotalCell' style:family='table-cell'>
              <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#DDDDDD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
            </style:style>
            <style:page-layout style:name='PL0'>
              <style:page-layout-properties fo:page-width='8.268in' fo:page-height='11.693in' style:print-orientation='portrait' fo:margin-top='1in' fo:margin-left='1.25in' fo:margin-bottom='1in' fo:margin-right='1.25in' style:num-format='1' style:writing-mode='lr-tb'>
                <style:footnote-sep style:width='0.007in' style:rel-width='33%' style:color='#000000' style:line-style='solid' style:adjustment='left'/>
              </style:page-layout-properties>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name='MP0' style:page-layout-name='PL0'/>
          </office:master-styles>";
    }
}
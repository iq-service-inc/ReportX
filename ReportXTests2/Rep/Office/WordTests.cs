using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Model;
using ReportXTests2;

namespace ReportX.Rep.Office.Tests
{

    [TestClass()]
    public class WordTests
    {
        [TestMethod()]
        public void rendrTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();

            Word report = new Word();
            report.setCol(data);
            if (cols != null && cols.Length > 0) report.changecut(cols);

            report.appendFullRow("測試增加一個 Header", null, "r-header-title");
            report.appendFullRow("測試客製化 Word", null, "r-header-secondary");
            report.appendFullRow("測試 CSS", null, "r-header-date");

            report.setCustomStyle(customOfficeCSS);

            // 測試寫入欄位
            ModelTR col = report.appendRow(report.cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";

            report.appendTable(data);

            string res = report.render();
            string fileName = sampleData.FileName + ".doc";
            ReportSaver.saveOfficeReport(fileName, res);

            Assert.IsNotNull(res);
        }

        const string customOfficeCSS = @"
            .r-header-title{
                font-size: 22px;
                font-weight: bold;
                text-align: center;
                background-color: #EAC !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-date{
                font-size: 18px;
                font-weight: bold;
                text-align: center;
                background-color: #CEA !important;
                -webkit-print-color-adjust: exact; 
            }
            .column{
                color: #FFF;
                text-align: center;
                background-color: #888 !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-secondary{
                color: #555;
                font-size: 14px;
                text-align: center;
                background-color: #FDE !important;
                -webkit-print-color-adjust: exact; 
            }";
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Office.Excel;
using ReportX.Rep.Office.Word;
using ReportX.Rep.OpenOffice.Ods;
using ReportX.Rep.OpenOffice.Odt;
using ReportXTests2;
using System;

namespace ReportX.Tests
{
    [TestClass()]
    public class ReportCreatorTest
    {
        [TestMethod()]
        public void renderExcelTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Excel> report = new ReportCreator<Excel>();
            string excel = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

            Assert.IsNotNull(excel);
            string fileName = sampleData.FileName + ".xls";
            ReportSaver.saveOfficeReport(fileName, excel);
        }
        [TestMethod()]
        public void renderExcelTest2()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<Excel> report = new ReportCreator<Excel>();
            string excel = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

            Assert.IsNotNull(excel);
            string fileName = sampleData.FileName + ".xls";
            ReportSaver.saveOfficeReport(fileName, excel);
        }

        [TestMethod()]
        public void renderWordTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Word> report = new ReportCreator<Word>();
            string word = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            Assert.IsNotNull(word);
            string fileName = sampleData.FileName + ".doc";
            ReportSaver.saveOfficeReport(fileName, word);
        }

        [TestMethod()]
        public void renderWordTest2()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<Word> report = new ReportCreator<Word>();
            string word = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            Assert.IsNotNull(word);
            string fileName = sampleData.FileName + ".doc";
            ReportSaver.saveOfficeReport(fileName, word);
        }

        [TestMethod()]
        public void renderOdtTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Odt> report = new ReportCreator<Odt>();
            string odt = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName + ".odt";
            ReportSaver.saveOpenOfficeReport(fileName, odt, report.report.meta);
        }

        [TestMethod()]
        public void renderOdtTest2()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<Odt> report = new ReportCreator<Odt>();
            string odt = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName + ".odt";
            ReportSaver.saveOpenOfficeReport(fileName, odt, report.report.meta);
        }

        [TestMethod()]
        public void renderOdsTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Ods> report = new ReportCreator<Ods>();
            string ods = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName + ".ods";
            ReportSaver.saveOpenOfficeReport(fileName, ods, report.report.meta);
        }

        [TestMethod()]
        public void renderOdsTest2()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<Ods> report = new ReportCreator<Ods>();
            string ods = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName + ".ods";
            ReportSaver.saveOpenOfficeReport(fileName, ods, report.report.meta);
        }
    }
}
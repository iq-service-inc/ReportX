using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Office;
using ReportX.Rep.OpenOffice;
using ReportXTests2;
using System;
using System.IO;

namespace ReportX.Tests
{
    [TestClass()]
    public class ReportFileTests
    {
        [TestMethod()]
        public void saveWordFileTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Word> report = new ReportCreator<Word>();
            report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string word = report.render();
            Assert.IsNotNull(word);
            string fileName = sampleData.FileName;

            ReportFile rf = new ReportFile(report.report);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }

        [TestMethod()]
        public void saveExcelFileTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Excel> report = new ReportCreator<Excel>();
            report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string excel = report.render();
            Assert.IsNotNull(excel);
            string fileName = sampleData.FileName;

            ReportFile rf = new ReportFile(report.report);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }

        [TestMethod()]
        public void saveOdtFileTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Odt> report = new ReportCreator<Odt>();
            report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName;

            ReportFile rf = new ReportFile(report.report);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }

        [TestMethod()]
        public void saveOdsFileTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string title = "測試資料";
            ReportCreator<Ods> report = new ReportCreator<Ods>();
            report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName;

            ReportFile rf = new ReportFile(report.report);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }

        [TestMethod()]
        public void saveMultiExcelCreatorFileTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            string fileName = sampleData.FileName;

            // 建立第一張 Excel
            ReportCreator<Excel> report1 = new ReportCreator<Excel>();
            report1.setInfo(data, cols, "第一個Excel", DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);


            // 建立第二張 Excel 
            ReportCreator<Excel> report2 = new ReportCreator<Excel>();
            report2.setInfo(data, cols, "第二個Excel", DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

            // 綁定兩個 Excel 
            MultiExcelBundler creator = new MultiExcelBundler();
            creator.addExcel(report1.report);
            creator.addExcel(report2.report);

            // 儲存成實體檔案
            ReportFile rf = new ReportFile(creator);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }
    }
}
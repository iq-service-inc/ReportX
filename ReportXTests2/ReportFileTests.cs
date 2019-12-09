using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX;
using ReportX.Rep.Office;
using ReportX.Rep.OpenOffice;
using ReportXTests2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            string word = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
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
            string excel = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

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
            string odt = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
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
            string ods = report.render(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string fileName = sampleData.FileName;

            ReportFile rf = new ReportFile(report.report);
            string path = rf.saveFile(fileName);
            Assert.IsTrue(File.Exists(path));
        }
    }
}
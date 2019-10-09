using Ionic.Zip;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportXTests2;
using ReportXTests2.Model;
using System;
using System.Data;
using System.IO;

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
            var dtData = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<ExcelReport> report = new ReportCreator<ExcelReport>();
            string excel = report.render<ExcelReport,ModelEmployeeTicket>(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            //string excel = report.render<ExcelReport>(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            Assert.IsNotNull(excel);
            if (File.Exists("creator.xls"))
            {
                File.Delete("creator.xls");
                File.AppendAllText("creator.xls", excel);
            }
            else
            {
                File.AppendAllText("creator.xls", excel);
            }
        }
        [TestMethod()]
        public void renderWordTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            var dtData = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<WordReport> report = new ReportCreator<WordReport>();
            //string word = report.render<WordReport,ModelEmployeeTicket>(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string word = report.render<WordReport>(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            Assert.IsNotNull(word);
            if (File.Exists("creator.doc"))
            {
                File.Delete("creator.doc");
                File.AppendAllText("creator.doc", word);
            }
            else
            {
                File.AppendAllText("creator.doc", word);
            }
        }
        [TestMethod()]
        public void renderOdtTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            var dtData = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<OdtReport> report = new ReportCreator<OdtReport>();
            string odt= report.render<OdtReport,ModelEmployeeTicket>(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            //string odt = report.render<OdtReport>(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string metaStr = report.renderOpenOfficeMeta();
            string dirPath = @".\META-INF";
            Assert.IsNotNull(metaStr);
            if (Directory.Exists(dirPath))
            {
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            else
            {
                Directory.CreateDirectory(dirPath);
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            Assert.IsNotNull(odt);
            if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", odt);
            }
            else
            {
                File.AppendAllText("content.xml", odt);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./creator.odt");
                }
            }
        }
        [TestMethod()]
        public void renderOdsTest()
        {
            SampleData sampleData = new SampleData();
            var cols = sampleData.ModelCol();
            var data = sampleData.ModelData();
            var dtData = sampleData.Dtdata();
            string title = "測試資料";
            ReportCreator<OdsReport> report = new ReportCreator<OdsReport>();
            string ods = report.render<OdsReport,ModelEmployeeTicket>(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            //string ods = report.render<OdsReport>(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string metaStr = report.renderOpenOfficeMeta();
            string dirPath = @".\META-INF";
            Assert.IsNotNull(metaStr);
            if (Directory.Exists(dirPath))
            {
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            else
            {
                Directory.CreateDirectory(dirPath);
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            Assert.IsNotNull(ods);
            if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", ods);
            }
            else
            {
                File.AppendAllText("content.xml", ods);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./creator.ods");
                }
            }
        }

    }
}
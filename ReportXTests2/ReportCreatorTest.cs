using Ionic.Zip;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Excel;
using ReportX.Rep.Integration;
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
            string[] cols = new string[4];
            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            ModelEmployeeTicket[] data = new ModelEmployeeTicket[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123",
                    data = "data",
                    tel = "0923456789"
                };
                data[i] = tmp;
            }
            string title = "測試資料";
            ReportCreator<ExcelReport> report = new ReportCreator<ExcelReport>();
            ReportCreator<ExcelReport> ex = report.ExcelReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string excel = ex.render();
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
            string[] cols = new string[4];
            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            ModelEmployeeTicket[] data = new ModelEmployeeTicket[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123",
                    data = "data",
                    tel = "0923456789"
                };
                data[i] = tmp;
            }
            string title = "測試資料";
            ReportCreator<WordReport> report = new ReportCreator<WordReport>();
            ReportCreator<WordReport> wd = report.WordReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            string word = wd.render();
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
            string[] cols = new string[4];
            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            ModelEmployeeTicket[] data = new ModelEmployeeTicket[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123",
                    data = "data",
                    tel = "0923456789"
                };
                data[i] = tmp;
            }
            string title = "測試資料";
            ReportCreator<OdtReport> report = new ReportCreator<OdtReport>();
            ReportCreator<OdtReport> odtr = report.OdtReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            report.CreateMeta("odt");
            var width = odtr.getColCount();
            string odt = odtr.render(width);
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
                string[] input = new string[2];
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
            string[] cols = new string[4];
            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            ModelEmployeeTicket[] data = new ModelEmployeeTicket[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123",
                    data = "data",
                    tel = "0923456789"
                };
                data[i] = tmp;
            }
            string title = "測試資料";
            ReportCreator<OdsReport> report = new ReportCreator<OdsReport>();
            ReportCreator<OdsReport> odsr = report.OdsReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
            report.CreateMeta("ods");
            string ods = odsr.render();
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
                string[] input = new string[2];
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
﻿using Ionic.Zip;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.OpenOffice.Odt;
using ReportXTests2;
using ReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            var dtData = sampleData.Dtdata();
            string title = "測試資料";
            string Creator = "測試人員";
            OdtReport report = new OdtReport(typeof(ModelEmployeeTicket));
            if (cols.Length > 0)
            {
                report.setcut(cols);
            }
            report.setTile(title);
            report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            report.setCreator(Creator);
            report.setCreatedDate();
            report.setColumn();

            //Model資料格式
            //report.setData(data);
            //report.setsum(data);

            //DateTime資料格式
            report.setData(dtData);
            report.setsum(dtData);
            report.CreateMeta("odt");
            var width = report.getColCount();
            var rpData = report.render(width);
            Assert.IsNotNull(rpData);
            if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", rpData);
            }
            else
            {
                File.AppendAllText("content.xml", rpData);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./report.odt");
                }
            }
            StreamReader str = new StreamReader(@"D:\ReportX\ReportXTests2\Sample\odt.txt");
            var ste = str.ReadToEnd();
            //此測試時，忽略時間資料，需註解掉 setDate,setCreateDate()
            Assert.AreEqual(rpData, ste);
        }
    }
}
﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Office.Excel;
using ReportXTests2;
using ReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Office.Excel.Tests
{
    [TestClass()]
    public class ExcelTests
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
            ExcelReport report = new ExcelReport(typeof(ModelEmployeeTicket));
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
            report.setData(data);
            report.setsum(data);

            //DateTime資料格式
            //report.setData(dtData);
            //report.setsum(dtData);

            var rpData =report.render(null);
            Assert.IsNotNull(rpData);
            if (File.Exists("report.xls"))
            {
                File.Delete("report.xls");
                File.AppendAllText("report.xls", rpData);
            }
            else
            {
                File.AppendAllText("report.xls", rpData);
            }
            //StreamReader str = new StreamReader(@"D:\ReportX\ReportXTests2\Sample\excel.txt");
            //var ste = str.ReadToEnd();
            //此測試時，忽略時間資料，需註解掉 setDate,setCreateDate()
            //Assert.AreEqual(rpData, ste);

        }
    }
}
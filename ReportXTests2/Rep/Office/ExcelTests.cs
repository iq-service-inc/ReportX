using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Office.Excel;
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
            string Creator = "測試人員";
            ReportCreator<Excel> report = new ReportCreator<Excel>(typeof(ModelEmployeeTicket));
            if (cols.Length > 0)
            {
                report.setcut(cols);
            }
            report.setTile(title);
            //report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            report.setCreator(Creator);
            //report.setCreatedDate();
            report.setColumn();
            report.setData(data);
            report.setsum(data, "Excel");
            var test =report.render();
            Assert.IsNotNull(test);
            if (File.Exists("report.xls"))
            {
                File.Delete("report.xls");
                File.AppendAllText("report.xls", test);
            }
            else
            {
                File.AppendAllText("report.xls", test);
            }
            StreamReader str = new StreamReader(@"D:\ReportX\ReportXTests2\Sample\excel.txt");
            var ste = str.ReadToEnd();
            Assert.AreEqual(test, ste);

        }
    }
}
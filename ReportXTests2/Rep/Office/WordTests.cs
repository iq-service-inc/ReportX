using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Office.Word;
using ReportXTests2;
using ReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Office.Word.Tests
{
    [TestClass()]
    public class WordTests
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
            WordReport report = new WordReport(typeof(ModelEmployeeTicket));
            if (cols.Length > 0)
            {
                report.setcut(cols);
            }
            report.setTile(title);
            //report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            report.setCreator(Creator);
            //report.setCreatedDate();
            report.setColumn();

            //Model資料格式
            //report.setData(data);
            //report.setsum(data);

            //DateTime資料格式
            report.setData(dtData);
            report.setsum(dtData);

            var rpData = report.render(null);
            Assert.IsNotNull(rpData);
            if (File.Exists("report.doc"))
            {
                File.Delete("report.doc");
                File.AppendAllText("report.doc", rpData);
            }
            else
            {
                File.AppendAllText("report.doc", rpData);
            }
            StreamReader str = new StreamReader(@"D:\ReportX\ReportXTests2\Sample\word.txt");
            var ste = str.ReadToEnd();
            //此測試時，忽略時間資料，需註解掉 setDate,setCreateDate()
            //Assert.AreEqual(rpData, ste);
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Excel;
using ReportX.Rep.Integration;
using ReportX.Rep.Odf;
using ReportXTests2.Model;
using System;
using System.IO;

namespace ReportX.Tests
{
    [TestClass()]
    public class ReportTests
    {
        [TestMethod()]
        public void excelResponse()
        {
            ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
            for (int i = 50 - 1; i >= 0; i--)
            {
                string s = Guid.NewGuid().ToString("N");
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123 ",
                    data = s,
                    tel = "0923456789"
                };
                data[i] = tmp;
            }

            string[] cols = new string[5];

            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            string title = "今日工事";

            Report Rpt = new Report();
            Excel excelRes = Rpt.excelResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘");
            string res = excelRes.render();
           
            if (File.Exists("data.xls"))
                File.Delete("data.xls");
            File.AppendAllText("data.xls", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);


        }
        [TestMethod()]
        public void odtResponse()
        {
            ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
            for (int i = 50 - 1; i >= 0; i--)
            {
                string s = Guid.NewGuid().ToString("N");
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123 ",
                    data = s,
                    tel = "0923456789"
                };
                data[i] = tmp;
            }

            string[] cols = new string[5];

            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            string title = "今日工事";

            Report Rpt = new Report();
            Odt odtRes = Rpt.OdtResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘");
            string res = odtRes.render();
            if (File.Exists("data.doc"))
                File.Delete("data.doc");
            File.AppendAllText("data.doc", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);


        }
        [TestMethod()]
        //綜合版測試
        public void FileReport()
        {
            ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
            for (int i = 50 - 1; i >= 0; i--)
            {
                string s = Guid.NewGuid().ToString("N");
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 100,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123 ",
                    data = "data" + i,
                    tel = "0923456789"+i
                };
                data[i] = tmp;
            }

            string[] cols = new string[5];

            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            string title = "今日工事";

            /*使用預設作法*/

            //Report rep = new Report();
            //FileReport file = rep.FileReport(data, cols, title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘", true);
            //string word = file.render(null, "word");
            //string excel = file.render(null, "excel");


            /*自定義作法*/

            FileReport file = new FileReport(typeof(ModelEmployeeTicket));
            file.setTile("標題");
            file.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            file.setCreatedDate();
            file.setColumn();
            file.setData(data);
            file.setsum(data);

            string word = file.render(null, "word");
            string excel = file.render(null, "excel");


            if (File.Exists("自定義綜合版.doc") && File.Exists("自定義綜合版.xls"))
            {
                File.Delete("自定義綜合版.doc");
                File.Delete("自定義綜合版.xls");

                File.AppendAllText("自定義綜合版.doc", word);
                File.AppendAllText("自定義綜合版.xls", excel);
            }
            else
            {
                File.AppendAllText("自定義綜合版.doc", word);
                File.AppendAllText("自定義綜合版.xls", excel);
            } 
        }
    }
}
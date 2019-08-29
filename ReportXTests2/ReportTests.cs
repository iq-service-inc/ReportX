using ICSharpCode.SharpZipLib.Zip;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Excel;
using ReportX.Rep.Integration;
using ReportX.Rep.Odf;
using ReportX.Rep.S5report;
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
            //excelRes.setCustomStyle();
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

            string[] cols = new string[6];

            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";
            string title = "今日工事";
            Report Rpt = new Report();
            Odt odtRes = Rpt.OdtResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            var width = odtRes.getColCount();
            string res = odtRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string outputFile = @".\result.odt";
                byte[] buffer = new byte[4096];
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                using (var input = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                using (var zip = new ZipOutputStream(output))
                {
                    ZipEntry entry = new ZipEntry(inputFile);
                    entry.DateTime = DateTime.Now;
                    zip.PutNextEntry(entry);
                    int readLength;
                    do
                    {
                        readLength = input.Read(buffer, 0, buffer.Length);
                        if (readLength > 0)
                        {
                            zip.Write(buffer, 0, readLength);
                        }
                    } while (readLength > 0);
                }
            }
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
                    tel = "0923456789" + i
                };
                data[i] = tmp;
            }

            string[] cols = new string[6];

            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";

            string title = "今日工事";

            /*使用預設作法*/

            Report rep = new Report();
            FileReport file = rep.FileReport(data, cols, title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘", true);
            string word = file.render(null, "word");
            string excel = file.render(null, "excel");


            /*自定義作法*/

            //FileReport file = new FileReport(typeof(ModelEmployeeTicket));
            //file.setTile("標題");
            //file.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            //file.setCreatedDate();
            //file.setColumn();
            //file.setData(data);
            //file.setsum(data);

            //string word = file.render(null, "word");
            //string excel = file.render(null, "excel");


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

        [TestMethod()]
        public void AmountReportTest()
        {
            Random r = new Random();
            var sum_correct = 0;
            var sum_wrong = 0;
            ModelKnowledgeAmount[] data = new ModelKnowledgeAmount[100];
            for (int i = 100 - 1; i >= 0; i--)
            {
                int num = r.Next(0, 30);
                int numw = r.Next(0, 5);
                string knowledge = "";
                string s = Guid.NewGuid().ToString("N");
                if (num%2==0)
                {
                    knowledge = "◎目錄" + i;
                }
                else
                {
                    knowledge = "目錄" + i;
                }
                ModelKnowledgeAmount tmp = new ModelKnowledgeAmount
                {
                    sequence = i + 1,
                    knowledge = knowledge,
                    correctAmount = num,
                    wrongAmount = numw,

                };
                data[i] = tmp;
            }
            
            var datetime = DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] cols = new string[4];
            cols[0] = "順序";
            cols[1] = "知識目錄";
            cols[2] = "有效知識";
            cols[3] = "無效知識";
            string title = "知識數量統計表";
            foreach (var item in data)
            {
                sum_correct += item.correctAmount;
                sum_wrong += item.wrongAmount;
            }
            Report Rpt = new Report();
            Amount AmountRes = Rpt.AmountReport(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), sum_correct, sum_wrong, "林家弘", true);
            var width = AmountRes.getColCount();
            string res = AmountRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string outputFile = @"./Amount(" + datetime + ").odt";
                byte[] buffer = new byte[4096];
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                using (var input = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                using (var zip = new ZipOutputStream(output))
                {
                    ZipEntry entry = new ZipEntry(inputFile);
                    entry.DateTime = DateTime.Now;
                    zip.PutNextEntry(entry);
                    int readLength;
                    do
                    {
                        readLength = input.Read(buffer, 0, buffer.Length);
                        if (readLength > 0)
                        {
                            zip.Write(buffer, 0, readLength);
                        }
                    } while (readLength > 0);
                }
            }
        }
    }
}
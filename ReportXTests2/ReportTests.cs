using ICSharpCode.SharpZipLib.Zip;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX.Rep.Excel;
using ReportX.Rep.Integration;
using ReportX.Rep.Odf;
using ReportX.Rep.S5report;
using ReportX.Rep.Word;
using ReportXTests2.Model;
using System;
using System.Data;
using System.IO;

namespace ReportX.Tests
{
    [TestClass()]
    public class ReportTests
    {
        [TestMethod()]
        public void excelResponse()
        {
            DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                                  new DataColumn("ID",typeof(int)),
                                  new DataColumn("標題",typeof(string)),
                                  new DataColumn("姓名",typeof(string)),
                                  new DataColumn("編號",typeof(decimal)),
                                  new DataColumn("資料",typeof(string)),
                                  new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                string a = Guid.NewGuid().ToString("N");
                row = dtTable.NewRow();
                row["ID"] = i;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = a.ToString();
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
            ModelEmployeeTicket[] data = new ModelEmployeeTicket[1];
            for (int i = 1 - 1; i >= 0; i--)
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

            //Excel excelRes = Rpt.excelResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);

            Excel excelRes = Rpt.excelResponse(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //excelRes.setCustomStyle();
            string res = excelRes.render(null);
            if (File.Exists("data.xls"))
                File.Delete("data.xls");
            File.AppendAllText("data.xls", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);


        }
        //  T[] 泛型陣列 & datatable輸入 odt 
        [TestMethod()]
        public void odtResponse()
        {

            DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                                  new DataColumn("ID",typeof(int)),
                                  new DataColumn("標題",typeof(string)),
                                  new DataColumn("姓名",typeof(string)),
                                  new DataColumn("編號",typeof(decimal)),
                                  new DataColumn("資料",typeof(string)),
                                  new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                string a = Guid.NewGuid().ToString("N");
                row = dtTable.NewRow();
                row["ID"] = i;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = a.ToString();
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
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
            cols[4] = "編號";
            string title = "今日工事";
            Report Rpt = new Report();
            //Odt odtRes = Rpt.OdtResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //dataTable 資料
           Odt odtRes = Rpt.OdtResponse(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            var width = odtRes.getColCount();
            string res = odtRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            Assert.IsNotNull(res);
            odtRes.CreateMeta("odt");
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @".\result.odt";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        // T[] 泛型陣列 & datatable輸入 ods         
        [TestMethod()]
        public void OdsResponseTest()
        {
            DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                                  new DataColumn("ID",typeof(int)),
                                  new DataColumn("標題",typeof(string)),
                                  new DataColumn("姓名",typeof(string)),
                                  new DataColumn("編號",typeof(decimal)),
                                  new DataColumn("資料",typeof(string)),
                                  new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                string a = Guid.NewGuid().ToString("N");
                row = dtTable.NewRow();
                row["ID"] = i;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = a.ToString();
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
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
            //Ods odsRes = Rpt.OdsResponse(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            Ods odsRes = Rpt.OdsResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            var width = odsRes.getColCount();
            string res = odsRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            odsRes.CreateMeta("ods");
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @".\result.ods";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        [TestMethod()]
        //綜合版測試 加入 datatable 資料輸入 (不額外寫 word test)
        public void FileReport()
        {
            DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                                  new DataColumn("ID",typeof(int)),
                                  new DataColumn("標題",typeof(string)),
                                  new DataColumn("姓名",typeof(string)),
                                  new DataColumn("編號",typeof(decimal)),
                                  new DataColumn("資料",typeof(string)),
                                  new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                string a = Guid.NewGuid().ToString("N");
                row = dtTable.NewRow();
                row["ID"] = i;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = a.ToString();
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
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
            //FileReport file = rep.FileReport(data, cols, title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘", true);
            FileReport file = rep.FileReport(dtTable, cols, title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘", true);
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

        //s5 報表 (無datatable格式輸入)
        [TestMethod()]
        public void AmountReportTest()
        {
            //生成假資料
            Random r = new Random();
            ModelKnowledgeAmount[] data = new ModelKnowledgeAmount[19];
            for (int i = 19 - 1; i >= 0; i--)
            {
                int num = r.Next(0, 30);
                int numw = r.Next(0, 5);
                string knowledge = "";
                string s = Guid.NewGuid().ToString("N");
                if (num % 2 == 0)
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
            //組成xml資料
            Report Rpt = new Report();
            Amount AmountRes = Rpt.AmountReport(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), "林家弘", true);
            string res = AmountRes.render();
            //產生xml
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            AmountRes.CreateMeta("odt");
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./Amount(" + datetime + ").odt";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        [TestMethod()]
        public void KBStaticReportTest()
        {
            Random r = new Random();
            int year = DateTime.Now.Year;//當前年
            int mouth = DateTime.Now.Month;//當前月
            int beforeYear = 0;
            int beforeMouth = 0;
            if (mouth <= 1)//如果當前月是一月，那麽年份就要減1
            {
                beforeYear = year - 1;
                beforeMouth = 12;//上個月
            }
            else
            {
                beforeYear = year;
                beforeMouth = mouth - 1;//上個月
            }
            string beforeMouthOneDay = beforeYear + "/" + beforeMouth + "/" + "1"; //上個月第一天
            string beforeMouthLastDay = beforeYear + "/" + beforeMouth + "/" + DateTime.DaysInMonth(year, beforeMouth); //上個月最後一天
            ModelKBStaticData[] data = new ModelKBStaticData[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                int num = r.Next(0, 30);
                int numw = r.Next(0, 5);
                string s = Guid.NewGuid().ToString("N");
                ModelKBStaticData tmp = new ModelKBStaticData
                {
                    number = i + 1,
                    knowledge = "目錄" + i,
                    knowledgeTitle = "標題" + i,
                    CreatedDate = DateTime.Now.ToString("yyyy/MM/dd hh:mm"),
                    Creater = "建立者" + i

                };
                data[i] = tmp;
            }

            var datetime = DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] cols = new string[5];
            cols[0] = "編號";
            cols[1] = "知識目錄";
            cols[2] = "知識標題";
            cols[3] = "建立時間";
            cols[4] = "建立人員";
            string title = "知識統計表";
            Report Rpt = new Report();
            KBStatic KBSRes = Rpt.KBStaticReport(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), beforeMouthOneDay, beforeMouthLastDay, "林家弘");
            var width = KBSRes.getColCount();
            string res = KBSRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            KBSRes.CreateMeta("odt");
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./KBStatics(" + datetime + ").odt";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        [TestMethod()]
        public void SignStatusReportTest()
        {
            Random r = new Random();
            int year = DateTime.Now.Year;//當前年
            int mouth = DateTime.Now.Month;//當前月
            int beforeYear = 0;
            int beforeMouth = 0;
            if (mouth <= 1)//如果當前月是一月，那麽年份就要減1
            {
                beforeYear = year - 1;
                beforeMouth = 12;//上個月
            }
            else
            {
                beforeYear = year;
                beforeMouth = mouth - 1;//上個月
            }
            string beforeMouthOneDay = beforeYear + "/" + beforeMouth + "/" + "1"; //上個月第一天
            string beforeMouthLastDay = beforeYear + "/" + beforeMouth + "/" + DateTime.DaysInMonth(year, beforeMouth); //上個月最後一天
            ModelSignStatusData[] data = new ModelSignStatusData[12];
            for (int i = 12 - 1; i >= 0; i--)
            {
                var addknowledgeReview = r.Next(0, 3);
                var addknowledgeBack = r.Next(0, 3);
                var addknowledgeWork = r.Next(0, 3);
                var deleteknowledgeReview = r.Next(0, 3);
                var deleteknowledgeBack = r.Next(0, 3);
                var deleteknowledgeWork = r.Next(0, 3);
                var updateknowledgeReview = r.Next(0, 3);
                var updateknowledgeBack = r.Next(0, 3);
                var updateknowledgeWork = r.Next(0, 3);
                var total = addknowledgeReview + addknowledgeBack + addknowledgeWork + deleteknowledgeReview + deleteknowledgeBack + deleteknowledgeWork
             + updateknowledgeReview + updateknowledgeBack + updateknowledgeWork;
                string s = Guid.NewGuid().ToString("N");
                ModelSignStatusData tmp = new ModelSignStatusData
                {
                    sequence = i + 1,
                    reviewDate = DateTime.Now.AddDays(i).ToString("yyyy/MM/dd"),
                    addknowledgeReview = addknowledgeReview,
                    addknowledgeBack = addknowledgeBack,
                    addknowledgeWork = addknowledgeWork,
                    deleteknowledgeReview = deleteknowledgeReview,
                    deleteknowledgeBack = deleteknowledgeBack,
                    deleteknowledgeWork = deleteknowledgeWork,
                    updateknowledgeReview = updateknowledgeReview,
                    updateknowledgeBack = updateknowledgeBack,
                    updateknowledgeWork = updateknowledgeWork,
                    total = total
                };
                data[i] = tmp;
            }

            var datetime = DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] cols = new string[15];
            cols[0] = "順序";
            cols[1] = "審核起始日";
            cols[2] = "新增知識";
            cols[3] = "修改知識";
            cols[4] = "刪除知識";
            cols[5] = "合計";
            cols[6] = "新增知識審核";
            cols[7] = "新增知識生效";
            cols[8] = "新增知識退件";
            cols[9] = "修改知識審核";
            cols[10] = "修改知識生效";
            cols[11] = "修改知識退件";
            cols[12] = "刪除知識審核";
            cols[13] = "刪除知識生效";
            cols[14] = "刪除知識退件";

            string title = "知識審核狀況統計表";
            Report Rpt = new Report();
            SignStatus SSRes = Rpt.SignStatusReport(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), beforeMouthOneDay, beforeMouthLastDay, "林家弘", true);
            var width = SSRes.getColCount();
            string res = SSRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            SSRes.CreateMeta("odt");
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./SignStatuss(" + datetime + ").odt";

                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        [TestMethod()]
        public void SignStatusReportOdsTest()
        {
            Random r = new Random();
            int year = DateTime.Now.Year;//當前年
            int mouth = DateTime.Now.Month;//當前月
            int beforeYear = 0;
            int beforeMouth = 0;
            if (mouth <= 1)//如果當前月是一月，那麽年份就要減1
            {
                beforeYear = year - 1;
                beforeMouth = 12;//上個月
            }
            else
            {
                beforeYear = year;
                beforeMouth = mouth - 1;//上個月
            }
            string beforeMouthOneDay = beforeYear + "/" + beforeMouth + "/" + "1"; //上個月第一天
            string beforeMouthLastDay = beforeYear + "/" + beforeMouth + "/" + DateTime.DaysInMonth(year, beforeMouth); //上個月最後一天
            ModelSignStatusData[] data = new ModelSignStatusData[12];
            for (int i = 12 - 1; i >= 0; i--)
            {
                var addknowledgeReview = r.Next(0, 3);
                var addknowledgeBack = r.Next(0, 3);
                var addknowledgeWork = r.Next(0, 3);
                var deleteknowledgeReview = r.Next(0, 3);
                var deleteknowledgeBack = r.Next(0, 3);
                var deleteknowledgeWork = r.Next(0, 3);
                var updateknowledgeReview = r.Next(0, 3);
                var updateknowledgeBack = r.Next(0, 3);
                var updateknowledgeWork = r.Next(0, 3);
                var total = addknowledgeReview + addknowledgeBack + addknowledgeWork + deleteknowledgeReview + deleteknowledgeBack + deleteknowledgeWork
             + updateknowledgeReview + updateknowledgeBack + updateknowledgeWork;
                string s = Guid.NewGuid().ToString("N");
                ModelSignStatusData tmp = new ModelSignStatusData
                {
                    sequence = i + 1,
                    reviewDate = DateTime.Now.AddDays(i).ToString("yyyy/MM/dd"),
                    addknowledgeReview = addknowledgeReview,
                    addknowledgeBack = addknowledgeBack,
                    addknowledgeWork = addknowledgeWork,
                    deleteknowledgeReview = deleteknowledgeReview,
                    deleteknowledgeBack = deleteknowledgeBack,
                    deleteknowledgeWork = deleteknowledgeWork,
                    updateknowledgeReview = updateknowledgeReview,
                    updateknowledgeBack = updateknowledgeBack,
                    updateknowledgeWork = updateknowledgeWork,
                    total = total
                };
                data[i] = tmp;
            }

            var datetime = DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] cols = new string[15];
            cols[0] = "順序";
            cols[1] = "審核起始日";
            cols[2] = "新增知識";
            cols[3] = "修改知識";
            cols[4] = "刪除知識";
            cols[5] = "合計";
            cols[6] = "新增知識審核";
            cols[7] = "新增知識生效";
            cols[8] = "新增知識退件";
            cols[9] = "修改知識審核";
            cols[10] = "修改知識生效";
            cols[11] = "修改知識退件";
            cols[12] = "刪除知識審核";
            cols[13] = "刪除知識生效";
            cols[14] = "刪除知識退件";

            string title = "知識審核狀況統計表";
            Report Rpt = new Report();
            SignStatusOds SSRes = Rpt.SignStatusReportOds(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), beforeMouthOneDay, beforeMouthLastDay, "林家弘", true);
            var width = SSRes.getColCount();
            string res = SSRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            SSRes.CreateMeta("ods");
            Assert.IsNotNull(res);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./SignStatuss(" + datetime + ").ods";

                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }
        [TestMethod()]
        public void KBStaticReportOdsTest()
        {
            Random r = new Random();
            int year = DateTime.Now.Year;//當前年
            int mouth = DateTime.Now.Month;//當前月
            int beforeYear = 0;
            int beforeMouth = 0;
            if (mouth <= 1)//如果當前月是一月，那麽年份就要減1
            {
                beforeYear = year - 1;
                beforeMouth = 12;//上個月
            }
            else
            {
                beforeYear = year;
                beforeMouth = mouth - 1;//上個月
            }
            string beforeMouthOneDay = beforeYear + "/" + beforeMouth + "/" + "1"; //上個月第一天
            string beforeMouthLastDay = beforeYear + "/" + beforeMouth + "/" + DateTime.DaysInMonth(year, beforeMouth); //上個月最後一天
            ModelKBStaticData[] data = new ModelKBStaticData[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                int num = r.Next(0, 30);
                int numw = r.Next(0, 5);
                string s = Guid.NewGuid().ToString("N");
                ModelKBStaticData tmp = new ModelKBStaticData
                {
                    number = i + 1,
                    knowledge = "目錄" + i,
                    knowledgeTitle = "標題" + i,
                    CreatedDate = DateTime.Now.ToString("yyyy/MM/dd hh:mm"),
                    Creater = "建立者" + i

                };
                data[i] = tmp;
            }

            var datetime = DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] cols = new string[5];
            cols[0] = "編號";
            cols[1] = "知識目錄";
            cols[2] = "知識標題";
            cols[3] = "建立時間";
            cols[4] = "建立人員";
            string title = "知識統計表";
            Report Rpt = new Report();
            KBStaticOds KBSRes = Rpt.KBStaticReportOds(data, cols, title, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), beforeMouthOneDay, beforeMouthLastDay, "林家弘");
            var width = KBSRes.getColCount();
            string res = KBSRes.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", res); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            KBSRes.CreateMeta("ods");

            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./KBStatics(" + datetime + ").ods";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }

        }
        [TestMethod()]
        public void ReportCreatorTest()
        {
            DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                                  new DataColumn("ID",typeof(int)),
                                  new DataColumn("標題",typeof(string)),
                                  new DataColumn("姓名",typeof(string)),
                                  new DataColumn("編號",typeof(decimal)),
                                  new DataColumn("資料",typeof(string)),
                                  new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                string a = Guid.NewGuid().ToString("N");
                row = dtTable.NewRow();
                row["ID"] = i;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = a.ToString();
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
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
            cols[4] = "標題";
            cols[5] = "編號";
            string title = "今日工事";
            Report res = new Report();
            //ReportCreator<WordReport> wd = new ReportCreator<WordReport>(typeof(ModelEmployeeTicket));
            ////file.setArray(typeof(ModelEmployeeTicket));
            //if (cols.Length > 0)
            //{
            //    wd.setcut(cols);
            //}
            //wd.setTile(title, "Word");
            //wd.setDate(DateTime.Now.AddDays(-1), DateTime.Now, "Word");
            //wd.setCreator(Creator, "Word");
            //wd.setCreatedDate("Word");
            //wd.setColumn();
            //wd.setData(data);
            //wd.setsum(data, "Word");
            ReportCreator<WordReport> wd = res.WordReport(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //ReportCreator<WordReport> wd = res.WordReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            string word = wd.render();
            //ReportCreator<ExcelReport> exc = new ReportCreator<ExcelReport>(typeof(ModelEmployeeTicket));
            //if (cols.Length > 0)
            //{
            //    exc.setcut(cols);
            //}
            //exc.setTile(title, "Excel");
            //exc.setDate(DateTime.Now.AddDays(-1), DateTime.Now, "Excel");
            //exc.setCreator(Creator, "Excel");
            //exc.setCreatedDate("Excel");
            //exc.setColumn();
            //exc.setData(data);
            //exc.setsum(data, "Excel");
            ReportCreator<ExcelReport> exc = res.ExcelReport(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //ReportCreator<ExcelReport> exc = res.ExcelReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            string excel = exc.render();
            if (File.Exists("creator.doc") && File.Exists("creator.xls"))
            {
                File.Delete("creator.doc");
                File.Delete("creator.xls");

                File.AppendAllText("creator.doc", word);
                File.AppendAllText("creator.xls", excel);
            }
            else
            {
                File.AppendAllText("creator.doc", word);
                File.AppendAllText("creator.xls", excel);
            }
            //ReportCreator<OdtReport> orp = new ReportCreator<OdtReport>(typeof(ModelEmployeeTicket));
            //if (cols.Length > 0)
            //{
            //    orp.setcut(cols);
            //}
            //orp.setTile(title,"Odt");
            //orp.setDate(DateTime.Now.AddDays(-1), DateTime.Now, "Odt");
            //orp.setCreator(Creator,"Odt");
            //orp.setCreatedDate("Odt");
            //orp.setColumn();
            //orp.setData(data);
            //orp.setsum(data,"Odt");
            ReportCreator<OdtReport> orp = res.OdtReport(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //ReportCreator<OdtReport> orp = res.OdtReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            var width = orp.getColCount();
            string odtRes = orp.render(width);
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", odtRes); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            orp.CreateMeta("odt");
            Assert.IsNotNull(odtRes);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./creator.odt";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
            //ReportCreator<OdsReport> osp = new ReportCreator<OdsReport>(typeof(ModelEmployeeTicket));
            //if (cols.Length > 0)
            //{
            //    osp.setcut(cols);
            //}
            //osp.setTile(title, "Odt");
            //osp.setDate(DateTime.Now.AddDays(-1), DateTime.Now, "Odt");
            //osp.setCreator(Creator, "Odt");
            //osp.setCreatedDate("Odt");
            //osp.setColumn();
            //osp.setData(data);
            //osp.setsum(data, "Odt");
            ReportCreator<OdsReport> osp = res.OdsReport(dtTable, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            //ReportCreator<OdsReport> osp = res.OdsReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
            string odsRes = osp.render();
            if (File.Exists("content.xml"))
                File.Delete("content.xml");
            File.AppendAllText("content.xml", odsRes); // 檔案存在 路徑: D:\CSharp\ReportX\ReportXTests2\bin\Debug
            orp.CreateMeta("ods");
            Assert.IsNotNull(odsRes);
            if (File.Exists("content.xml"))
            {
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @"./creator.ods";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {

                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            }
        }

    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportX;
using ReportX.Rep.Excel;
using ReportX.Rep.Word;
using ReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i,
                    posttitle = "測試"+i,
                    name = "SOL" + i,
                    number ="000"+i
                };
                data[i] = tmp;
            }

            var Rpt = new Report();
            Excel excelRes = Rpt.excelResponse(data, "測試title", Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘");
            // act  
            // assert  
            string res = excelRes.render();
            Trace.WriteLine(res);


            if (File.Exists("data.xls")) File.Delete("data.xls");
            File.AppendAllText("data.xls", res);
            Assert.IsNotNull(res);
        }

        [TestMethod()]
        public void WordResponse()
        {

            ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
            for (int i = 50 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i,
                    posttitle = "測試" + i,
                    name = "SOL" + i,
                    number = "000" + i
                };
                data[i] = tmp;
            }

            Report sdf = new Report();
            WordReport wr =sdf.WordResponse(data, "測試title", Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "林家弘");

            string s = wr.render();
            

            if (File.Exists("data.doc")) File.Delete("data.doc");
            Trace.WriteLine(s);

            File.AppendAllText("data.doc", s);
            Assert.IsNotNull(s);

   
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyReportX.Rep.Excel;
using MyReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyReportX.Rep.Excel.Tests
{
    [TestClass()]
    public class ExcelReportTests
    {
        [TestMethod()]
        public void formatDate()
        {
            ExcelReport erp = new ExcelReport(typeof(ModelEmployeeTicket));
            string res = erp.formatDate(Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"));
            Trace.WriteLine(res);
            Assert.AreEqual("2017/01/20 - 2017/01/20", res);
        }
    }
}
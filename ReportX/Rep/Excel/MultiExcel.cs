using ReportX.Rep.Model;
using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportX.Rep.Excel
{
    public class MultiExcel
    {
        private List<ModelMultiExcel> ExcelReportlist;

        public MultiExcel(List<ExcelReport> list)
        {
            ExcelReportlist= list.Select(x=> new ModelMultiExcel { report=x ,cid= Guid.NewGuid().ToString() }).ToList();
        }

        public string render()
        {
            ViewMultiExcel report = new ViewMultiExcel(ExcelReportlist);
            return report.render();
        }
    }
}

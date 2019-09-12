using ReportX.Rep.View.S5View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
    public class ModelSignStatusOds
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string datetime { get; set; }
        public string sheetName { get; set; }
        public string dateRange { get; set; }
        public ViewStyleSignStatusOds style { get; set; }
        public ViewBodySignStatusOds body { get; set; }
        public ModelSignStatusOds()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
        }
    }
}

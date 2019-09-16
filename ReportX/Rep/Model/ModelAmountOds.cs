using ReportX.Rep.View.S5View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
    public class ModelAmountOds
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string datetime { get; set; }
        public string sheetName { get; set; }
        public ViewStyleAmountOds style { get; set; }
        public ViewBodyAmountOds body { get; set; }
        public ModelAmountOds()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
        }
    }
}

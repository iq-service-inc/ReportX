using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
   public class ModelOds
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string sheetName { get; set; }
        public ViewStyleOds style { get; set; }
        public ViewBodyOds body { get; set; }


        public ModelOds()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
        }
    }
}

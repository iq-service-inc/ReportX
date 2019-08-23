using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
    public  class ModelOdt
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string sheetName { get; set; }
        public ViewStyleOdt style { get; set; }
        public ViewBodyOdt body { get; set; }


        public ModelOdt()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
        }
    }
}

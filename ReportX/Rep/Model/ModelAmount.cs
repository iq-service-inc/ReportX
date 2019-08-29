using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
   public class ModelAmount
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string datetime { get; set; }
        public string sheetName { get; set; }
        public bool mark { get; set; }
        public ViewStyleAmount style { get; set; }
        public ViewBodyAmount body { get; set; }
        public ModelAmount()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
            mark = false;
        }
    }
}

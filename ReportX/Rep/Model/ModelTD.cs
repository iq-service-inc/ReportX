using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
    public class ModelTD
    {
        public object data { get; set; }
        public int colspan { get; set; }
        public int rowspan { get; set; }
        public string fontSize { get; set; }
        public string align { get; set; }
        public bool bold { get; set; }
        public string bgcolor { get; set; }
        public string style { get; set; }
        public string className { get; set; }

        public ModelTD()
        {
            colspan = 1;
            rowspan = 1;
            bold = false;
        }
    }
}

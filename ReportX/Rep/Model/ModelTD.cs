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
        public string col { get; set; } //欄位
        public string sum_c { get; set; } //有效知識總和
        public string sum_w { get; set; } //無效知識總和   

        public ModelTD()
        {
            colspan = 1;
            rowspan = 1;
            bold = false;
        }
    }
}

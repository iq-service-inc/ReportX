using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Model
{
    public class ModelWord
    {
        public int colNum { get; set; }
        public string author { get; set; }
        public string company { get; set; }
        public string sheetName { get; set; }
        public ViewStyle style { get; set; }
        public ViewBody body { get; set; }


        public ModelWord()
        {
            company = "IQ-data";
            sheetName = "DownloadWord";
        }

    }
}

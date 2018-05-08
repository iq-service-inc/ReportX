using MyReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyReportX.Rep.Model
{
        public class ModelExcel
        {
            public int colNum { get; set; }
            public string author { get; set; }
            public string company { get; set; }
            public string sheetName { get; set; }
            public ViewStyle style { get; set; }
            public ViewBody body { get; set; }


            public ModelExcel()
            {
                company = "IQ-data";
                sheetName = "DownloadExcel";
            }

        }

    
}

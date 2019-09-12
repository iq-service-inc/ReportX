using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewMultiExcel
    {
        private List<ModelMultiExcel> m;

        public ViewMultiExcel(List<ModelMultiExcel> model)
        {
            m = model;
        }

        public string render()
        {
            string worksheet = String.Join("",m.Select(x=> string.Format(Worksheettemplate, x.report.getsheetName(),x.cid)).ToList()),
                   bodyrender = String.Join("", m.Select(x => string.Format(bodytemplate, x.cid,x.report.render(null))).ToList()); 

            return string.Format(template, worksheet, bodyrender);

        }
        string Worksheettemplate = @"
                <x:ExcelWorksheet>
                <x:Name>{0}</x:Name>
                <x:WorksheetSource HRef='cid:{1}'/>
                </x:ExcelWorksheet>
            ";
//必須要靠左 排版會影響輸出 
        string bodytemplate = @"
---=BOUNDARY_EXCEL
Content-ID: {0}
Content-Type: text/html; charset='utf-8'
{1}
";

 //MIME必須要靠左，排版會影響輸出 
        string template = @"MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary=""-=BOUNDARY_EXCEL""

---=BOUNDARY_EXCEL
Content-Type: text/html; charset=""utf-8""

            <html xmlns:o=""urn:schemas-microsoft-com:office:office""
            xmlns:x=""urn:schemas-microsoft-com:office:excel"">
            <head>
            <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
            <xml>
            <o:DocumentProperties>
            </o:DocumentProperties>
            </xml>
            <xml>
            <x:ExcelWorkbook>
            <x:ExcelWorksheets>
            {0}
            </x:ExcelWorksheets>
            </x:ExcelWorkbook>
            </xml>     
            </head>
            </html>

            {1}
---=BOUNDARY_EXCEL--
";
    }

}

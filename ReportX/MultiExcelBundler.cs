using ReportX.Rep.Office;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReportX
{
    /// <summary>
    /// 多個 Excel 合成一個專用工具
    /// </summary>
    public class MultiExcelBundler
    {
        private List<Excel> reports;

        public MultiExcelBundler()
        {
            reports = new List<Excel>();
        }

        /// <summary>
        /// 添加 Excel 報表，如果使用 ReportCreator 可以把 ReportCreator.report 這個變數加入
        /// </summary>
        /// <param name="report">Excel 報表</param>
        public void addExcel(Excel report)
        {
            if (report != null)
                reports.Add(report);
        }

        /// <summary>
        /// 將多個 Excel 綁定成一個，並生成新的內容字串
        /// </summary>
        /// <param name="width">寬度</param>
        /// <returns>Excel 內容字串，存檔成 .xls 就可以使用</returns>
        public string render(int? width = null)
        {
            if (reports.Count < 1) return null;

            StringBuilder sheetTabs = new StringBuilder();
            StringBuilder sheetContent = new StringBuilder();
            foreach (Excel rep in reports)
            {
                string id = Guid.NewGuid().ToString();
                sheetTabs.Append(string.Format(sheetTabsTemplate, rep.getsheetName(), id));
                sheetContent.Append(string.Format(sheetsTempalte, id, rep.render(width)));
            }
            string res = string.Format(multiExcelTemplate, sheetTabs.ToString(), sheetContent.ToString());
            return res;
        }


        const string sheetsTempalte =
@"---=BOUNDARY_EXCEL
Content-ID: {0}
Content-Type: text/html; charset='utf-8'
{1}
";

        const string sheetTabsTemplate =
@"<x:ExcelWorksheet>
<x:Name>{0}</x:Name>
<x:WorksheetSource HRef='cid:{1}'/>
</x:ExcelWorksheet>
";


        const string multiExcelTemplate =
@"MIME-Version: 1.0
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
            <x:ExcelWorksheets>{0}</x:ExcelWorksheets>
            </x:ExcelWorkbook>
            </xml>     
            </head>
            </html>
{1}
---=BOUNDARY_EXCEL--
";

    }
}

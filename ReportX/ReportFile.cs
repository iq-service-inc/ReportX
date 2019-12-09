using Ionic.Zip;
using ReportX.Rep.Common;
using ReportX.Rep.Office;
using ReportX.Rep.OpenOffice;
using System;
using System.IO;

namespace ReportX
{
    /// <summary>
    /// 將報表字串存檔成實體檔案專用工具，支援 IReport 與 MultiExcelCreator 兩種規格的檔案產生
    /// </summary>
    public class ReportFile
    {
        private IReportX report;
        private MultiExcelBundler excel_creator;

        /// <summary>
        /// 報表隔離區路徑
        /// </summary>
        public string isolatedPath { get; private set; }

        /// <summary>
        /// 報表名稱(包含附檔名)
        /// </summary>
        public string fileName { get; private set; }

        /// <summary>
        /// 單一報表檔案建立專用
        /// </summary>
        /// <param name="report"></param>
        public ReportFile(IReportX report)
        {
            if (report == null) throw new Exception("Report Object is null");
            this.report = report;
            isolatedPath = createIsolated();
        }

        /// <summary>
        /// 複數 Excel 報表合成一個專用
        /// </summary>
        /// <param name="excel_creator"></param>
        public ReportFile(MultiExcelBundler excel_creator)
        {
            if (excel_creator == null) throw new Exception("MultiExcelCreator is null");
            this.excel_creator = excel_creator;
            isolatedPath = createIsolated();
        }

        /// <summary>
        /// 將報表儲存成實體檔案，並回傳儲存路徑
        /// </summary>
        /// <param name="fileName">報表名稱(不用副檔名)</param>
        /// <param name="width">寬度</param>
        /// <returns>報表儲存路徑</returns>
        public string saveFile(string name, int? width = null)
        {
            fileName = string.IsNullOrEmpty(name) ? Guid.NewGuid().ToString() : name;
            string path = "";
            if (excel_creator != null)
            {
                fileName = $"{name}.xls";
                path = $"{isolatedPath}\\{fileName}";
                string content = excel_creator.render(width);
                saveOfficeReport(path, content);
            }
            else
            {
                string file_ext = getFileExtensionName();
                fileName = $"{name}{file_ext}";
                path = $"{isolatedPath}\\{fileName}";
                string content = report.render(width);
                if (report is AbsOpenOffice) saveOpenOfficeReport(path, content);
                else saveOfficeReport(path, content);
            }
            return path;
        }

        /// <summary>
        /// 如果報表已經不需要再使用，則可以呼叫此方法刪除檔案，否則需要自行刪除
        /// </summary>
        public void deleteReportFile()
        {
            if (string.IsNullOrEmpty(isolatedPath)) return;
            if (Directory.Exists(isolatedPath))
            {
                Directory.Delete(isolatedPath, true);
            }
        }

        /// <summary>
        /// 建立隔離區(目錄)避免在建立報表的過程中與其他報表產生的暫存檔案產生衝突  
        /// e.g. content.xml (OpenOffice)
        /// </summary>
        /// <returns></returns>
        private string createIsolated()
        {
            string isolatedName = Guid.NewGuid().ToString();
            while (Directory.Exists(isolatedName)) isolatedName = Guid.NewGuid().ToString();
            Directory.CreateDirectory(isolatedName);
            return isolatedName;
        }


        private string getFileExtensionName()
        {
            if (report is Excel) return ".xls";
            else if (report is Word) return ".doc";
            else if (report is Ods) return ".ods";
            else if (report is Odt) return ".odt";
            
            else return "";
        }


        private void saveOfficeReport(string fileName, string content)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
                File.AppendAllText(fileName, content);
            }
            else
            {
                File.AppendAllText(fileName, content);
            }
        }

        private void saveOpenOfficeReport(string fileName, string content)
        {
            string metaDirPath = $"{isolatedPath}\\META-INF";
            string metaFilePath = $"{metaDirPath}\\manifest.xml";
            string contentFilePath = $"{isolatedPath}\\content.xml";
            string metaStr = "";

            if (report is Ods) metaStr = ((Ods)report).meta;
            else if (report is Odt) metaStr = ((Odt)report).meta;

            if (!Directory.Exists(metaDirPath))
                Directory.CreateDirectory(metaDirPath);

            if (File.Exists(metaFilePath)) File.Delete(metaFilePath);
            if (File.Exists(contentFilePath)) File.Delete(contentFilePath);

            File.AppendAllText(metaFilePath, metaStr);
            File.AppendAllText(contentFilePath, content);

            using (var zip = new ZipFile(""))
            {
                zip.AddFile(contentFilePath, "\\");
                zip.AddFile(metaFilePath, "\\META-INF");
                zip.Save(fileName);
            }
        }
    }
}

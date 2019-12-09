using Ionic.Zip;
using ReportX.Rep.Common;
using ReportX.Rep.Office;
using ReportX.Rep.OpenOffice;
using System;
using System.IO;

namespace ReportX
{
    public class ReportFile
    {
        private IReportX report;

        /// <summary>
        /// 報表隔離區路徑
        /// </summary>
        public string isolatedPath { get; private set; }

        /// <summary>
        /// 報表名稱(包含附檔名)
        /// </summary>
        public string fileName { get; private set; }


        public ReportFile(IReportX report)
        {
            if (report == null) throw new Exception("Report Object is null");
            this.report = report;
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
            isolatedPath = createIsolated();
            string file_ext = getFileExtensionName();
            fileName = $"{name}.{file_ext}";
            string path = $"{isolatedPath}\\{fileName}";
            string content = report.render(width);
            if (report is AbsOpenOffice) saveOpenOfficeReport(path, content);
            else saveOfficeReport(path, content);
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


        private string createIsolated()
        {
            string isolatedName = Guid.NewGuid().ToString();
            while (Directory.Exists(isolatedName)) isolatedName = Guid.NewGuid().ToString();
            Directory.CreateDirectory(isolatedName);
            return isolatedName;
        }

        private string getFileExtensionName()
        {
            if (report is Excel) return "xls";
            else if (report is Word) return "doc";
            else if (report is Ods) return "ods";
            else if (report is Odt) return "odt";
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
            string dirPath = @".\META-INF";
            string metaStr = "";

            if (report is Ods) metaStr = ((Ods)report).meta;
            else if (report is Odt) metaStr = ((Odt)report).meta;

            if (Directory.Exists(dirPath))
            {
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            else
            {
                Directory.CreateDirectory(dirPath);
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", metaStr);
            }
            if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", content);
            }
            else
            {
                File.AppendAllText("content.xml", content);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(fileName);
                }
            }
        }
    }
}

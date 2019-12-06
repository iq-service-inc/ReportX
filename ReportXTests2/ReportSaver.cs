using Ionic.Zip;
using System.IO;

namespace ReportXTests2
{
    public static class ReportSaver
    {
        public static void saveOfficeReport(string fileName, string content)
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

        public static void saveOpenOfficeReport(string fileName, string content, string metaStr)
        {
            string dirPath = @".\META-INF";

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

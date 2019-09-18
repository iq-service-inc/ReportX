using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ReportX
{

    public class ReportCreator<T> where T : IReportX
    {
        static T report { get; set; }


        public MemberInfo[] modeli;
        protected Type type { get; set; }
        protected DataTable data { get; set; }
        protected string[] oldcols { get; set; }
        protected string[] newcols { get; set; }
        protected List<ModelTR> trs { get; }
        public static string[] cols { get; set; }
        public string sheetName { get; set; }
        private ModelExcel excel;
        private ModelWord word;
        private ModelOdt odt;

        public ReportCreator(Type type)
        {


            List<MemberInfo> list_cols = new List<MemberInfo>();
            modeli = type.GetMembers();
            foreach (var member in type.GetMembers())
            {
                Present attr = member.GetCustomAttribute<Present>();
                if (attr == null) continue;

                int MetadataToken = member.MetadataToken,
                    inserted_index = 0;

                // sory by MetadataToken (declaration)
                for (int i = 0; i < list_cols.Count; i++)
                {
                    inserted_index = i;
                    if (MetadataToken < list_cols[i].MetadataToken) break;
                    inserted_index = i + 1;
                }
                list_cols.Insert(inserted_index, member);
            }
            string[] str_cols = new string[list_cols.Count]; //取得標題數量

            for (int i = 0; i < list_cols.Count; i++)
                str_cols[i] = list_cols[i].GetCustomAttribute<Present>().getName();//取得標題名稱

            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            report = (T)Activator.CreateInstance(typeof(T), type);
        }
        public string render(int? width = null)
        {
         
            return report.render(width);
        }


        public void setDate(DateTime from, DateTime? to = null, string type = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");
            switch (type)
            {
                case "Word":
                    report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
                    break;
                case "Excel":
                    report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
                    break;
                case "Odt":
                    report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
                    break;
                case "Ods":
                    report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
                    break;
                default:
                    break;
            }
        }
        public void setCreator(string creator, string type = null)
        {
            report.setData(author: creator);
            switch (type)
            {
                case "Word":
                    report.appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
                    break;
                case "Excel":
                    report.appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
                    break;
                case "Odt":
                    report.appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");
                    break;
                case "Ods":
                    report.appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");
                    break;
                default:
                    break;
            }
        }
        public void setTile(string title, string type = null)
        {
            report.setData(sheetName: title);
            switch (type)
            {
                case "Word":
                    report.appendFullRow(title, null, "r-header-title");
                    break;
                case "Excel":
                    report.appendFullRow(title, null, "r-header-title");
                    break;
                case "Odt":
                    report.appendFullRow(title, "TableCellData", "Title");
                    break;
                case "Ods":
                    report.appendFullRow(title, "TableCellData", "Title");
                    break;
                default:
                    break;
            }

        }
        public void setCreatedDate(string type = null)
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            switch (type)
            {
                case "Word":
                    report.appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
                    break;
                case "Excel":
                    report.appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
                    break;
                case "Odt":
                    report.appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
                    break;
                case "Ods":
                    report.appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
                    break;
                default:
                    break;
            }
        }

        public void setColumn()
        {
            var intersectResult = oldcols.Intersect(newcols);
            cols = intersectResult.ToArray();
            ModelTR col = report.appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }


        public void setData<T>(T[] data)
        {
            report.appendTable(data);
        }
        public void setsum<T>(T[] data, string type) //總筆數
        {
            string lastRowStyle = "";
            string lastClassName = "";
            switch (type)
            {
                case "Word":
                    lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
                    report.appendRow(new { value = "總筆數", colspan = report.getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數
                    break;
                case "Excel":
                    lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
                    report.appendRow(new { value = "總筆數", colspan = report.getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數
                    break;
                case "Odt":
                    lastRowStyle = "TotalCell"; //預設CSS
                    lastClassName = "Word";
                    report.appendRow(new { value = data.Length, colspan = report.getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數                    break;
                    break;
                case "Ods":
                    lastRowStyle = "TotalCell"; //預設CSS
                    lastClassName = "Word";
                    report.appendRow(new { value = data.Length, colspan = report.getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數
                    break;
                default:
                    break;
            }

        }
        // 傳入欲顯示欄位標題 之陣列
        public void setcut(string[] cut)
        {
            newcols = cut;
            report.changecut(cut);
        }
        public void CreateMeta(string type)
        {
            var str = "";
            switch (type)
            {
                case "odt":
                    str = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.text'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='settings.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";
                    break;
                case "ods":
                    str = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.spreadsheet'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";
                    break;
                default:
                    break;
            }
            string dirPath = @".\META-INF";
            if (Directory.Exists(dirPath))
            {
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", str);
            }
            else
            {
                Directory.CreateDirectory(dirPath);
                if (File.Exists("META-INF/manifest.xml"))
                    File.Delete("META-INF/manifest.xml");
                File.AppendAllText("META-INF/manifest.xml", str);
            }
        }
        public int getColCount()
        {
            return cols.Length;
        }
    }

}
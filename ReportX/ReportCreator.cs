using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.Office.Excel;
using ReportX.Rep.Office.Word;
using ReportX.Rep.OpenOffice.Ods;
using ReportX.Rep.OpenOffice.Odt;
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
        public string[] oldcols { get; set; }
        public string[] newcols { get; set; }
        protected List<ModelTR> trs { get; }
        public static string[] cols { get; set; }
        public string sheetName { get; set; }

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
        public ReportCreator(DataTable data)
        {
            trs = new List<ModelTR>();
            string[] str_cols = new string[data.Columns.Count];

            for (int i = 0; i < data.Columns.Count; i++)
                str_cols[i] = data.Columns[i].ToString();
            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            report = (T)Activator.CreateInstance(typeof(T), data);
        }
        public ReportCreator()
        {
        }

        public string render(int? width = null)
        {
            return report.render(width);
        }
        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;
            string classname = typeof(T).Name;
            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");
            if (typeof(ExcelReport).Name == classname || typeof(WordReport).Name == classname)
                report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
            else if (typeof(OdtReport).Name == classname || typeof(OdsReport).Name == classname)
                report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
        }
        public void setCreator(string creator)
        {
            string classname = typeof(T).Name;
            report.setData(author: creator);
            if (typeof(ExcelReport).Name == classname || typeof(WordReport).Name == classname)
                report.appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
            else if (typeof(OdtReport).Name == classname || typeof(OdsReport).Name == classname)
                report.appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");


        }
        public void setTile(string title)
        {
            string classname = typeof(T).Name;
            report.setData(sheetName: title);
            if (typeof(ExcelReport).Name == classname || typeof(WordReport).Name == classname)
                report.appendFullRow(title, null, "r-header-title");
            else if (typeof(OdtReport).Name == classname || typeof(OdsReport).Name == classname)
                report.appendFullRow(title, "TableCellData", "Title");

        }
        public void setCreatedDate()
        {
            string classname = typeof(T).Name;
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            if (typeof(ExcelReport).Name == classname || typeof(WordReport).Name == classname)
                report.appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
            else if (typeof(OdtReport).Name == classname || typeof(OdsReport).Name == classname)
                report.appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
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
        public void setData(DataTable data)
        {
            report.appendTable(data);
        }
        public void setsum<T>(T[] data, Type type) //總筆數
        {
            string lastRowStyle = "";
            string lastClassName = "";
            if (typeof(ExcelReport).Name == type.Name || typeof(WordReport).Name == type.Name)
            {
                lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
                report.appendRow(new { value = "總筆數", colspan = report.getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數
            }
            else if (typeof(OdtReport).Name == type.Name || typeof(OdsReport).Name == type.Name)
            {
                lastRowStyle = "TotalCell"; //預設CSS
                lastClassName = "Word";
                report.appendRow(new { value = data.Length, colspan = report.getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數 
            }
        }
        public void setsum(DataTable data) //總筆數
        {
            string classname = typeof(T).Name;
            string lastRowStyle = "";
            string lastClassName = "";

            if (typeof(ExcelReport).Name == classname || typeof(WordReport).Name == classname)
            {
                lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
                report.appendRow(new { value = "總筆數", colspan = report.getColCount() - 1, style = lastRowStyle }, data.Select().Count());//統計資料數
            }
            else if (typeof(OdtReport).Name == classname || typeof(OdsReport).Name == classname)
            {
                lastRowStyle = "TotalCell"; //預設CSS
                lastClassName = "Word";
                report.appendRow(new { value = data.Select().Count(), colspan = report.getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數
            }
        }
        // 傳入欲顯示欄位標題 之陣列
        public void setcut(string[] cut)
        {
            newcols = cut;
            report.changecut(cut);
        }
        public string CreateMeta(Type type)
        {
            var classname = type.Name;
            var str = "";
            if (typeof(OdtReport).Name == classname)
                str = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.text'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='settings.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";
            if (typeof(OdsReport).Name == classname)
                str = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'><manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.spreadsheet'/><manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml'/><manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml'/></manifest:manifest>";
            return str;
        }
        public int getColCount()
        {
            return cols.Length;
        }
        public string render<T,S>(S[] data, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false) where T : IReportX
        {
            ReportCreator<T> report = new ReportCreator<T>(typeof(S));
            string res = "";
            if (cols.Length > 0)
            {
                report.setcut(cols);
            }

            report.setTile(title);
            report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            report.setCreator(Creator);
            report.setCreatedDate();
            report.setColumn();
            report.setData(data);
            if (end) //如果要顯示結算筆數 end =true;
            {
                report.setsum(data,typeof(T));
            }
            if (typeof(T) == typeof(OdtReport) || typeof(T) == typeof(Odt))
            {
                int width = report.getColCount();
                res = report.render(width);
            }
            else
            {
                res = report.render();

            }
            return res;
        }
        public string render<T>(DataTable data, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false) where T : IReportX
        {
            ReportCreator<T> report = new ReportCreator<T>(data);
            string res = "";
            if (cols.Length > 0)
            {
                report.setcut(cols);
            }

            report.setTile(title);
            report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
            report.setCreator(Creator);
            report.setCreatedDate();
            report.setColumn();
            report.setData(data);
            if (end) //如果要顯示結算筆數 end =true;
            {
                report.setsum(data);
            }
            if (typeof(T) == typeof(OdtReport) || typeof(T) == typeof(Odt))
            {
                int width = report.getColCount();
                res = report.render(width);
            }
            else
            {
                res = report.render();

            }
            return res;
        }
        public string renderOpenOfficeMeta()
        {
            ReportCreator<AbsOpenOffice> meta = new ReportCreator<AbsOpenOffice>();
            string metaStr = meta.CreateMeta(typeof(T));
            return metaStr;
        }
    }
    public class ExcelReport : Excel
    {
        string customCSS = @"
            .r-header-title{
                font-size: 24px;
                font-weight: bold;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-date{
                font-size: 20px;
                font-weight: bold;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }
            .column{
                color: #FFF;
                text-align: center;
                background-color: #555 !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-secondary{
                color: #555;
                font-size: 14px;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }             
        ";


        public ExcelReport(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public ExcelReport(DataTable model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        // 設定製表標題
        public void setTile(string title)
        {
            setData(sheetName: title);
            appendFullRow(title, null, "r-header-title");
        }

        // 設定製表日期 : 帶入參數 yyyy/MM/dd yyyy/MM/dd
        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");

            appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
        }

        // 設定製表人
        public void setCreator(string creator)
        {
            setData(author: creator);
            appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
        }

        // 設定製表時間 :取得現在時間
        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
        }

        // 設定資料欄位
        public void setColumn()
        {
            ModelTR col = appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }

        // 塞入資料
        public void setData<T>(T[] data)
        {
            appendTable(data);
        }
        public void setData(DataTable data)
        {
            appendTable(data);
        }
        // 傳入欲顯示欄位標題 之陣列
        public void setcut(string[] cut)
        {
            changecut(cut);
        }

        public void setsum<T>(T[] data) //總筆數欄位
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            appendRow(new { value = "總筆數", colspan = getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數

        }
        public void setsum(DataTable data) //總筆數欄位
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            appendRow(new { value = "總筆數", colspan = getColCount() - 1, style = lastRowStyle }, data.Select().Count());//統計資料數

        }
    }
    public class WordReport : Word
    {
        string customCSS = @"
            .r-header-title{
                font-size: 24px;
                font-weight: bold;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-date{
                font-size: 20px;
                font-weight: bold;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }
            .column{
                color: #FFF;
                text-align: center;
                background-color: #555 !important;
                -webkit-print-color-adjust: exact; 
            }
            .r-header-secondary{
                color: #555;
                font-size: 14px;
                text-align: center;
                background-color: #DEF !important;
                -webkit-print-color-adjust: exact; 
            }             
        ";

        public WordReport(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public WordReport(DataTable model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public void setTile(string title)
        {
            setData(sheetName: title);
            appendFullRow(title, null, "r-header-title");
        }

        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");

            appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
        }

        public void setCreator(string creator)
        {
            setData(author: creator);
            appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
        }

        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
        }

        public void setColumn()
        {
            ModelTR col = appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }

        public void setData<T>(T[] data)
        {
            appendTable(data);
        }
        public void setData(DataTable data)
        {
            appendTable(data);
        }
        // 傳入欲顯示欄位標題 之陣列
        public void setcut(string[] cut)
        {
            changecut(cut);
        }

        public void setsum<T>(T[] data) //總筆數
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            appendRow(new { value = "總筆數", colspan = getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數

        }
        public void setsum(DataTable data) //總筆數
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            appendRow(new { value = "總筆數", colspan = getColCount() - 1, style = lastRowStyle }, data.Select().Count());//統計資料數

        }
    }
    public class OdsReport : Ods
    {
        string customCSS = @" <office:automatic-styles>
    <style:style style:name='ColumnWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#555555' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties fo:color='#FFFFFF' style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
    </style:style>
    <style:style style:name='Word' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap'/>
      <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
    </style:style>
    <style:style style:name='TotalWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDDDDD'/>
      <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
    </style:style>
    <style:style style:name='TitleWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold'/>
    </style:style>
    <style:style style:name='DateRangeWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='15pt' style:font-size-asian='15pt' style:font-size-complex='15pt' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold'/>
    </style:style>
    <style:style style:name='CreaterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border='2pt solid #AAAAAA' style:vertical-align='middle' fo:wrap-option='wrap' fo:background-color='#DDEEFF' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties fo:color='#555555' style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體' fo:font-size='11pt' style:font-size-asian='11pt' style:font-size-complex='11pt'/>
    </style:style>
    <style:style style:name='TableColumn' style:family='table-column'>
      <style:table-column-properties />
    </style:style>
    <style:style style:name='TableRow' style:family='table-row'>
      <style:table-row-properties style:row-height='auto' style:use-optimal-row-height='false' fo:break-before='auto'/>
    </style:style>
    <style:style style:name='ta1' style:family='table' style:master-page-name='mp1'>
      <style:table-properties table:display='true' style:writing-mode='lr-tb'/>
    </style:style>
    <style:page-layout style:name='pm1'>
      <style:page-layout-properties fo:margin-top='0.5in' fo:margin-bottom='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' style:print-orientation='portrait' style:print-page-order='ttb' style:first-page-number='continue' style:scale-to='100%' style:table-centering='none' style:print='objects charts drawings'/>
      <style:header-style>
        <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-bottom='0in'/>
      </style:header-style>
      <style:footer-style>
        <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-top='0in'/>
      </style:footer-style>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name='mp1' style:page-layout-name='pm1'>
      <style:header/>
      <style:header-left style:display='false'/>
      <style:footer/>
      <style:footer-left style:display='false'/>
    </style:master-page>
  </office:master-styles>
        ";

        public OdsReport(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public OdsReport(DataTable model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public void setTile(string title)
        {
            setData(sheetName: title);
            appendFullRow(title, "TableCellData", "Title");
        }
        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");

            appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
        }
        public void setCreator(string creator)
        {
            setData(author: creator);
            appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");
        }
        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
        }
        public void setColumn()
        {
            ModelTR col = appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }
        public void setData<T>(T[] data)
        {
            appendTable(data);
        }
        public void setData(DataTable data)
        {
            appendTable(data);
        }
        public void setcut(string[] cut)
        {
            changecut(cut);
        }
        public void setsum<T>(T[] data) //總筆數
        {
            string lastRowStyle = "TotalCell"; //預設CSS
            string lastClassName = "Word";
            appendRow(new { value = data.Length, colspan = getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數

        }
        public void setsum(DataTable data) //總筆數
        {
            string lastRowStyle = "TotalCell"; //預設CSS
            string lastClassName = "Word";
            appendRow(new { value = data.Select().Count(), colspan = getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數

        }
    }
    public class OdtReport : Odt
    {

        string customCSS = @"
    <office:automatic-styles>
        <style:style style:name='TableColumn' style:family='table-column'>
          <style:table-column-properties style:column-width='auto'/>
        </style:style>
        <style:style style:name='Table' style:family='table' style:master-page-name='MP0'>
          <style:table-properties  fo:margin-left='0in' table:align='center'/>
        </style:style>
        <style:style style:name='TableRow' style:family='table-row'>
          <style:table-row-properties/>
        </style:style>
        <style:style style:name='TableCellData' style:family='table-cell'>
          <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#DDEEFF' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
        </style:style>
        <style:style style:name='Title' style:parent-style-name='內文' style:family='paragraph'>
          <style:paragraph-properties fo:widows='2' fo:orphans='2' fo:break-before='page' fo:text-align='center'/>
        </style:style>
        <style:style style:name='TitleWord' style:parent-style-name='預設段落字型' style:family='text'>
          <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt'/>
        </style:style>
        <style:style style:name='TitleDateWord' style:parent-style-name='內文' style:family='paragraph'>
          <style:paragraph-properties fo:text-align='center'/>
          <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:font-weight='bold' style:font-weight-asian='bold' style:font-weight-complex='bold' fo:font-size='15pt' style:font-size-asian='15pt' style:font-size-complex='15pt'/>
        </style:style>
        <style:style style:name='TitleTimeWord' style:parent-style-name='內文' style:family='paragraph'>
          <style:paragraph-properties fo:text-align='center'/>
          <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:color='#555555' fo:font-size='10.5pt' style:font-size-asian='10.5pt' style:font-size-complex='10.5pt'/>
        </style:style>
        <style:style style:name='TitleCell' style:family='table-cell'>
          <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#555555' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
        </style:style>
        <style:style style:name='TitleCellWord' style:parent-style-name='內文' style:family='paragraph'>
          <style:paragraph-properties fo:text-align='center'/>
          <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' fo:color='#FFFFFF'/>
        </style:style>
        <style:style style:name='CellWord' style:family='table-cell'>
          <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
        </style:style>
        <style:style style:name='Word' style:parent-style-name='內文' style:family='paragraph'>
          <style:text-properties style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體'/>
        </style:style>
        <style:style style:name='TotalCell' style:family='table-cell'>
          <style:table-cell-properties fo:border='0.0104in solid #AAAAAA' fo:background-color='#DDDDDD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0208in' fo:padding-bottom='0in' fo:padding-right='0.0208in'/>
        </style:style>
        <style:page-layout style:name='PL0'>
          <style:page-layout-properties fo:page-width='8.268in' fo:page-height='11.693in' style:print-orientation='portrait' fo:margin-top='1in' fo:margin-left='1.25in' fo:margin-bottom='1in' fo:margin-right='1.25in' style:num-format='1' style:writing-mode='lr-tb'>
            <style:footnote-sep style:width='0.007in' style:rel-width='33%' style:color='#000000' style:line-style='solid' style:adjustment='left'/>
          </style:page-layout-properties>
        </style:page-layout>
      </office:automatic-styles>
      <office:master-styles>
        <style:master-page style:name='MP0' style:page-layout-name='PL0'/>
      </office:master-styles>           
        ";

        public OdtReport(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public OdtReport(DataTable model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public void setTile(string title)
        {
            setData(sheetName: title);
            appendFullRow(title, "TableCellData", "Title");
        }
        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");

            appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
        }
        public void setCreator(string creator)
        {
            setData(author: creator);
            appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");
        }
        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
        }
        public void setColumn()
        {
            ModelTR col = appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }
        public void setData<T>(T[] data)
        {
            appendTable(data);
        }
        public void setData(DataTable data)
        {
            appendTable(data);
        }
        public void setcut(string[] cut)
        {
            changecut(cut);
        }
        public void setsum<T>(T[] data) //總筆數
        {
            string lastRowStyle = "TotalCell"; //預設CSS
            string lastClassName = "Word";
            appendRow(new { value = data.Length, colspan = getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數

        }
        public void setsum(DataTable data) //總筆數
        {
            string lastRowStyle = "TotalCell"; //預設CSS
            string lastClassName = "Word";
            appendRow(new { value = data.Select().Count(), colspan = getColCount() - 1, style = lastRowStyle, className = lastClassName });//統計資料數

        }
    }
}
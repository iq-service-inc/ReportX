using ReportX.Rep.Model;
using System;

namespace ReportX.Rep.Excel
{
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

        // 設定製表標題
        public void setTile(string title)
        {
            setExcel(sheetName: title);
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
            setExcel(author: creator);
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
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportX.Rep.Model;

namespace ReportX.Rep.Word
{
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
        public void setsum<T>(DataTable data) //總筆數
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            appendRow(new { value = "總筆數", colspan = getColCount() - 1, style = lastRowStyle }, data.Select().Count());//統計資料數

        }
    }

}

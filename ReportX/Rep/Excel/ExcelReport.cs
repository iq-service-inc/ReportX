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


        public void setTile(string title)
        {
            setExcel(sheetName: title);
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
            setExcel(author: creator);
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


    }
}

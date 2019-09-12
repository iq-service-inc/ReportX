using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Odf
{
    public class OdsReport : Ods
    {
        string customCSS = @"  <office:automatic-styles>
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
            setOds(sheetName: title);
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
            setOds(author: creator);
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
        // 傳入欲顯示欄位標題 之陣列
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

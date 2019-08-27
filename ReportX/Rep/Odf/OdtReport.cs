using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Odf
{
   public class OdtReport:Odt
    {
       
 string customCSS = @"


<office:automatic-styles>
    <style:style style:name='TableColumn' style:family='table-column'>
      <style:table-column-properties />
    </style:style>
    <style:style style:name='TableColumn' style:family='table-column'>
      <style:table-column-properties />
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

        public void setTile(string title)
        {
            setWord(sheetName: title);
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
            setWord(author: creator);
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

    }
}


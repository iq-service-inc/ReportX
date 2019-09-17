using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class KBStaticReportOds:KBStaticOds
    {
        string customCSS = @"  <office:automatic-styles>
    <style:style style:name='CountWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
      <style:text-properties style:font-name='Times New Roman' style:font-name-asian='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='12pt' style:font-size-asian='12pt' style:font-size-complex='12pt'/>
    </style:style>
    <style:style style:name='DataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
    </style:style>
    <style:style style:name='DataCenterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N30'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='ColumnFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='thin solid #000000' fo:border-right='none'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='ColumnCenterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='none' fo:border-right='none'/>
    </style:style>
    <style:style style:name='ColumnEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='none' fo:border-right='thin solid #000000'/>
    </style:style>
    <style:style style:name='FirstDataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='none' fo:border-left='thin solid #000000' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='EndDataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='none' fo:border-left='none' fo:border-right='thin solid #000000' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TotalFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='thin solid #000000' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TotalWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='none' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
    </style:style>
    <style:style style:name='TotalCenterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N30'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='none' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TotalEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='none' fo:border-right='thin solid #000000' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='Notice' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
      <style:text-properties fo:font-size='10pt' style:font-size-asian='10pt' style:font-size-complex='10pt'/>
    </style:style>
    <style:style style:name='TitleWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='標楷體' style:font-name-asian='標楷體' style:font-name-complex='標楷體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt' style:font-family-generic='script'/>
    </style:style>
    >
    <style:style style:name='DateRange' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='Times New Roman' style:font-name-asian='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='14pt' style:font-size-asian='14pt' style:font-size-complex='14pt'/>
    </style:style>
    <style:style style:name='User' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='Times New Roman' style:font-name-asian='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='8pt' style:font-size-asian='8pt' style:font-size-complex='8pt'/>
    </style:style>
    <style:style style:name='Column' style:family='table-column'>
      <style:table-column-properties  style:column-width='auto'/>
    </style:style>
   <style:style style:name='Row' style:family='table-row'>
      <style:table-row-properties style:row-height='auto'/>
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
  </office:master-styles>";
        public KBStaticReportOds(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public void setTile(string title)
        {
            setData(sheetName: title);
        }

        public void setCreator(string creator)
        {
            setData(author: creator);
        }
        public void setCreatedDayRange(string firstday, string lastday)
        {
            setData(dateRange: firstday + " - " + lastday);
        }
        public void setCreatedDate(string dateTime)
        {
            setData(dateTime: dateTime);
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
            string lastClassName = "Data";
            appendRow(new { colspan = getColCount() - 2, style = lastRowStyle, className = lastClassName, value = data });//統計資料數

        }
    }
}

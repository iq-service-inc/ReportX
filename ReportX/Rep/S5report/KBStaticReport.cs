using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class KBStaticReport : KBStatic
    {
        string customCSS = @"<office:automatic-styles>
    <style:style style:name='Title' style:parent-style-name='內文' style:master-page-name='MP0' style:family='paragraph'>
      <style:paragraph-properties fo:break-before='page' fo:text-align='center' fo:margin-bottom='0.1111in'/>
      <style:text-properties style:font-name='標楷體' style:font-name-asian='標楷體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt'/>
    </style:style>
    <style:style style:name='DateRange' style:parent-style-name='內文Web' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='center' fo:margin-bottom='0.0555in' fo:margin-left='0in' fo:margin-right='0in'>
        <style:tab-stops/>
      </style:paragraph-properties>
      <style:text-properties style:font-name='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='14pt' style:font-size-asian='14pt' style:font-size-complex='14pt'/>
    </style:style>
    <style:style style:name='User' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties style:line-break='normal' fo:text-align='end'/>
      <style:text-properties fo:font-size='8pt' style:font-size-asian='8pt' style:font-size-complex='8pt'/>
    </style:style>
    <style:style style:name='TopNotice' style:parent-style-name='內文Web' style:family='paragraph'>
      <style:paragraph-properties fo:margin-bottom='0.0555in' fo:margin-left='0in' fo:margin-right='0in'>
        <style:tab-stops/>
      </style:paragraph-properties>
      <style:text-properties style:font-name='Times New Roman' style:font-name-complex='Times New Roman'/>
    </style:style>
    <style:style style:name='TableColumn6' style:family='table-column'>
      <style:table-column-properties style:column-width='0.3534in'/>
    </style:style>
    <style:style style:name='TableColumn7' style:family='table-column'>
      <style:table-column-properties style:column-width='2.9513in'/>
    </style:style>
    <style:style style:name='TableColumn8' style:family='table-column'>
      <style:table-column-properties style:column-width='2.0437in'/>
    </style:style>
    <style:style style:name='TableColumn9' style:family='table-column'>
      <style:table-column-properties style:column-width='0.7826in'/>
    </style:style>
    <style:style style:name='TableColumn10' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5506in'/>
    </style:style>
    <style:style style:name='Table' style:family='table'>
      <style:table-properties style:width='6.6819in' style:rel-width='100%' fo:margin-left='0in' table:align='center'/>
    </style:style>
    <style:style style:name='TableRow' style:family='table-row'>
      <style:table-row-properties/>
    </style:style>
    <style:style style:name='TableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' fo:background-color='#EAF1FD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.075in' fo:padding-bottom='0in' fo:padding-right='0.075in'/>
    </style:style>
    <style:style style:name='Data' style:parent-style-name='內文' style:family='paragraph'>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體' fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt' fo:hyphenate='false'/>
    </style:style>
    <style:style style:name='DataTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='CenterWord' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體' fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt' fo:hyphenate='false'/>
    </style:style>
    <style:style style:name='DownNotice' style:parent-style-name='內文Web' style:family='paragraph'>
      <style:paragraph-properties fo:margin-left='0in' fo:margin-right='0'>
        <style:tab-stops/>
      </style:paragraph-properties>
      <style:text-properties fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt'/>
    </style:style>
      <style:page-layout style:name='PL0'>
      <style:page-layout-properties fo:page-width='8.268in' fo:page-height='11.693in' style:print-orientation='portrait' fo:margin-top='0.7875in' fo:margin-left='0.7875in' fo:margin-bottom='0.7875in' fo:margin-right='0.7875in' style:num-format='1' style:writing-mode='lr-tb'>
        <style:footnote-sep style:width='0.007in' style:rel-width='33%' style:color='#000000' style:line-style='solid' style:adjustment='left'/>
      </style:page-layout-properties>
    </style:page-layout>
  </office:automatic-styles>
    <office:master-styles>
    <style:master-page style:name='MP0' style:page-layout-name='PL0'/>
  </office:master-styles>";
        public KBStaticReport(Type model) : base(model)
        {
            setCustomStyle(customCSS);
        }
        public void setTile(string title)
        {
            setTitle(sheetName: title);
        }

        public void setCreator(string creator)
        {
            setTitle(author: creator);
        }
        public void setCreatedDayRange(string firstday,string lastday)
        {
            setTitle(dateRange: firstday+" - "+lastday);
        }
        public void setCreatedDate(string dateTime)
        {
            setTitle(dateTime: dateTime);
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

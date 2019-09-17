using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class AmountReportOds:AmountOds
    {
        string customCSS = @"  <office:automatic-styles>
    <style:style style:name='Column' style:family='table-column'>
      <style:table-column-properties fo:break-before='auto' style:column-width='auto'/>
    </style:style>
    <style:style style:name='Row' style:family='table-row'>
      <style:table-row-properties style:row-height='auto' fo:break-before='auto' style:use-optimal-row-height='false'/>
    </style:style>
    <style:style style:name='ta1' style:family='table' style:master-page-name='mp1'>
      <style:table-properties table:display='true' style:writing-mode='lr-tb'/>
    </style:style>
    <style:style style:name='TableName' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:text-align-source='fix' style:repeat-content='false' fo:wrap-option='wrap' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='標楷體' fo:font-size='18pt' style:font-name-asian='標楷體1' style:font-size-asian='18pt' style:font-name-complex='標楷體1' style:font-size-complex='18pt'/>
    </style:style>
    <style:style style:name='User' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:text-align-source='fix' style:repeat-content='false' fo:wrap-option='wrap' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='Times New Roman' fo:font-size='8pt' style:font-name-asian='Times New Roman' style:font-size-asian='8pt' style:font-name-complex='Times New Roman' style:font-size-complex='8pt'/>
    </style:style>
    <style:style style:name='ColumneFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='none' style:text-align-source='fix' style:repeat-content='false' fo:border-left='0.002cm solid #000000' fo:border-right='none' fo:border-top='0.002cm solid #000000' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='DataFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='none' style:text-align-source='fix' style:repeat-content='false' fo:border-left='0.002cm solid #000000' fo:border-right='none' fo:border-top='none' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TotalFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='0.002cm solid #000000' style:text-align-source='fix' style:repeat-content='false' fo:border-left='0.002cm solid #000000' fo:border-right='none' fo:border-top='none' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='ColumnWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='none' style:text-align-source='fix' style:repeat-content='false' fo:border-left='none' fo:border-right='none' fo:border-top='0.002cm solid #000000' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='DataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:text-align-source='fix' style:repeat-content='false' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
    </style:style>
    <style:style style:name='DataMarkWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:text-align-source='fix' style:repeat-content='false' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
      <style:text-properties fo:color='#ff0000'/>
    </style:style>
    <style:style style:name='CorrectWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:text-align-source='fix' style:repeat-content='false' fo:padding-bottom='0.035cm' fo:padding-left='0.035cm' fo:padding-right='0cm' fo:padding-top='0.035cm' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='end'/>
    </style:style>
    <style:style style:name='TotalCorrectWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='0.002cm solid #000000' style:text-align-source='fix' style:repeat-content='false' fo:border-left='none' fo:padding-bottom='0.035cm' fo:padding-left='0.035cm' fo:padding-right='0cm' fo:padding-top='0.035cm' fo:border-right='none' fo:border-top='none' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='end'/>
    </style:style>
    <style:style style:name='ColumnEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='none' style:text-align-source='fix' style:repeat-content='false' fo:border-left='none' fo:border-right='0.002cm solid #000000' fo:border-top='0.002cm solid #000000' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='DataEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='none' style:text-align-source='fix' style:repeat-content='false' fo:border-left='none' fo:padding-bottom='0.035cm' fo:padding-left='0.035cm' fo:padding-right='0cm' fo:padding-top='0.035cm' fo:border-right='0.002cm solid #000000' fo:border-top='none' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='end'/>
    </style:style>
    <style:style style:name='TotalWrongWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-bottom='0.002cm solid #000000' style:text-align-source='fix' style:repeat-content='false' fo:border-left='none' fo:padding-bottom='0.035cm' fo:padding-left='0.035cm' fo:padding-right='0cm' fo:padding-top='0.035cm' fo:border-right='0.002cm solid #000000' fo:border-top='none' style:vertical-align='middle'/>
      <style:paragraph-properties fo:text-align='end'/>
    </style:style>
    <style:page-layout style:name='Mpm3'>
      <style:page-layout-properties style:num-format='1' style:print-orientation='portrait' fo:margin-top='1.27cm' fo:margin-bottom='1.27cm' fo:margin-left='1.905cm' fo:margin-right='1.905cm' style:print-page-order='ttb' style:first-page-number='continue' style:scale-to='100%' style:print='charts drawings objects'/>
      <style:header-style>
        <style:header-footer-properties fo:min-height='1.27cm' fo:margin-left='1.905cm' fo:margin-right='1.905cm' fo:margin-bottom='0cm'/>
      </style:header-style>
      <style:footer-style>
        <style:header-footer-properties fo:min-height='1.27cm' fo:margin-left='1.905cm' fo:margin-right='1.905cm' fo:margin-top='0cm'/>
      </style:footer-style>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name='mp1' style:page-layout-name='Mpm3'>
      <style:header/>
      <style:header-left style:display='false'/>
      <style:footer/>
      <style:footer-left style:display='false'/>
    </style:master-page>
  </office:master-styles>";
        public AmountReportOds(Type model) : base(model)
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
            int sum_correct = 0;
            int sum_wrong = 0;
            string lastRowStyle = "TotalCell"; //預設CSS
            string lastClassName = "Total";
            foreach (T item in data)
            {
                foreach (var prop in item.GetType().GetProperties())
                {
                    switch (prop.Name)
                    {
                        case "correctAmount":
                            sum_correct += (int)prop.GetValue(item, null);
                            break;
                        case "wrongAmount":
                            sum_wrong += (int)prop.GetValue(item, null);
                            break;
                        default:
                            break;
                    }
                }
            }
            appendRow(new { colspan = getColCount() - 2, style = lastRowStyle, className = lastClassName, value = "合計" },
                new { value = sum_correct, style = "correct", className = lastClassName },
                new { value = sum_wrong, style = "wrong", className = lastClassName });//統計資料數

        }
    }
}

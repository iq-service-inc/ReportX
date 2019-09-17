﻿using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class AmountReport : Amount
    {
        string customCSS = @"<office:automatic-styles>
    <style:style style:name='Title' style:parent-style-name='內文' style:master-page-name='MP0' style:family='paragraph'>
      <style:paragraph-properties fo:break-before='page' fo:text-align='center' fo:margin-bottom='0.1111in'/>
      <style:text-properties style:font-name='標楷體' style:font-name-asian='標楷體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt'/>
    </style:style>
    <style:style style:name='User' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties style:line-break='normal' fo:text-align='end'/>
      <style:text-properties fo:font-size='8pt' style:font-size-asian='8pt' style:font-size-complex='8pt'/>
    </style:style>
    <style:style style:name='TableColumn4' style:family='table-column'>
      <style:table-column-properties style:column-width='0.543in'/>
    </style:style>
    <style:style style:name='TableColumn5' style:family='table-column'>
      <style:table-column-properties style:column-width='4.2069in'/>
    </style:style>
    <style:style style:name='TableColumn6' style:family='table-column'>
      <style:table-column-properties style:column-width='0.968in'/>
    </style:style>
    <style:style style:name='TableColumn7' style:family='table-column'>
      <style:table-column-properties style:column-width='0.9638in'/>
    </style:style>
    <style:style style:name='Table' style:family='table'>
      <style:table-properties style:width='6.6819in' fo:margin-left='0in' table:align='center'/>
    </style:style>
    <style:style style:name='TitleTableRow' style:family='table-row'>
      <style:table-row-properties/>
    </style:style>
    <style:style style:name='TableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' fo:background-color='#EAF1FD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.075in' fo:padding-bottom='0in' fo:padding-right='0.075in'/>
    </style:style>
    <style:style style:name='ContentTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' fo:background-color='#EAF1FD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='Data' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體'/>
    </style:style>
    <style:style style:name='SequenceDataTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='ContentDataTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.0416in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='KnowledgeDataTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0.0416in'/>
    </style:style>
    <style:style style:name='KnowledgeData' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0.1666in'/>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體'/>
    </style:style>
    <style:style style:name='MarkContentData' style:parent-style-name='內文' style:family='paragraph'>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體' fo:color='#FF0000'/>
    </style:style>
    <style:style style:name='ContentData' style:parent-style-name='內文' style:family='paragraph'>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體'/>
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
        public AmountReport(Type model) : base(model)
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
            string lastClassName = "Data";
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
                new {   value = sum_correct,  className = lastClassName },
                new {   value = sum_wrong,  className = lastClassName });//統計資料數

        }
    }
}

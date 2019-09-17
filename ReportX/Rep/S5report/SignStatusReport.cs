using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class SignStatusReport : SignStatus
    {
        string customCSS = @"  <office:automatic-styles>
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
    <style:style style:name='TableColumn5' style:family='table-column'>
      <style:table-column-properties style:column-width='0.4326in'/>
    </style:style>
    <style:style style:name='TableColumn6' style:family='table-column'>
      <style:table-column-properties style:column-width='0.8055in'/>
    </style:style>
    <style:style style:name='TableColumn7' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5416in'/>
    </style:style>
    <style:style style:name='TableColumn8' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5368in'/>
    </style:style>
    <style:style style:name='TableColumn9' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5416in'/>
    </style:style>
    <style:style style:name='TableColumn10' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5416in'/>
    </style:style>
    <style:style style:name='TableColumn11' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5368in'/>
    </style:style>
    <style:style style:name='TableColumn12' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5416in'/>
    </style:style>
    <style:style style:name='TableColumn13' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5416in'/>
    </style:style>
    <style:style style:name='TableColumn14' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5354in'/>
    </style:style>
    <style:style style:name='TableColumn15' style:family='table-column'>
      <style:table-column-properties style:column-width='0.5423in'/>
    </style:style>
    <style:style style:name='TableColumn16' style:family='table-column'>
      <style:table-column-properties style:column-width='0.584in'/>
    </style:style>
    <style:style style:name='Table' style:family='table'>
      <style:table-properties style:width='6.6819in' fo:margin-left='0in' table:align='center'/>
    </style:style>
    <style:style style:name='HeaderTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' fo:background-color='#EAF1FD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0.075in' fo:padding-bottom='0in' fo:padding-right='0.075in'/>
    </style:style>
    <style:style style:name='Word' style:parent-style-name='內文' style:family='paragraph'>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體' fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt' fo:hyphenate='false'/>
    </style:style>
    <style:style style:name='DataTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='CenterWord' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體' fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt' fo:hyphenate='false'/>
    </style:style>
    <style:style style:name='TableRow' style:family='table-row'>
      <style:table-row-properties/>
    </style:style>
    <style:style style:name='FooterTableCell' style:family='table-cell'>
      <style:table-cell-properties fo:border='0.0104in outset #000000' fo:background-color='#EAF1FD' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
    </style:style>
    <style:style style:name='EndWord' style:parent-style-name='內文' style:family='paragraph'>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0.0416in'/>
      <style:text-properties style:font-name='新細明體' style:font-name-complex='新細明體'  fo:font-size='9pt' style:font-size-asian='9pt' style:font-size-complex='9pt' fo:hyphenate='false'/>
    </style:style>
    <style:style style:name='TableColumn102' style:family='table-column'>
      <style:table-column-properties style:column-width='6.693in'/>
    </style:style>
    <style:style style:name='NoticeTable' style:family='table'>
      <style:table-properties style:width='6.693in' fo:margin-left='0in' table:align='center'/>
    </style:style>
    <style:style style:name='NoticeTableRow' style:family='table-cell'>
      <style:table-cell-properties fo:border='none' style:writing-mode='lr-tb' style:vertical-align='middle' fo:padding-top='0in' fo:padding-left='0in' fo:padding-bottom='0in' fo:padding-right='0in'/>
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
        public SignStatusReport(Type model) : base(model)
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
        public void setSecondColumn()
        {
            ModelTR col = appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "secondColumn";
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
            int Sum_addReview = 0;
            int Sum_addBack = 0;
            int Sum_addWork = 0;
            int Sum_updateReview = 0;
            int Sum_updateBack = 0;
            int Sum_updateWork = 0;
            int Sum_deleteReview = 0;
            int Sum_deleteBack = 0;
            int Sum_deleteWork = 0;
            foreach (T item in data)
            {
                foreach (var prop in item.GetType().GetProperties())
                {
                    switch (prop.Name)
                    {
                        case "addknowledgeReview":
                            Sum_addReview += (int)prop.GetValue(item, null);
                            break;
                        case "addknowledgeWork":
                            Sum_addWork += (int)prop.GetValue(item, null);
                            break;
                        case "addknowledgeBack":
                            Sum_addBack += (int)prop.GetValue(item, null);
                            break;
                        case "updateknowledgeReview":
                            Sum_updateReview += (int)prop.GetValue(item, null);
                            break;
                        case "updateknowledgeWork":
                            Sum_updateWork += (int)prop.GetValue(item, null);
                            break;
                        case "updateknowledgeBack":
                            Sum_updateBack += (int)prop.GetValue(item, null);
                            break;
                        case "deleteknowledgeReview":
                            Sum_deleteReview += (int)prop.GetValue(item, null);
                            break;
                        case "deleteknowledgeWork":
                            Sum_deleteWork += (int)prop.GetValue(item, null);
                            break;
                        case "deleteknowledgeBack":
                            Sum_deleteBack += (int)prop.GetValue(item, null);
                            break;
                        default:
                            break;
                    }
                }
            }
            var sum_total = Sum_addReview + Sum_addWork + Sum_addBack + Sum_updateReview + Sum_updateWork + Sum_updateBack
                + Sum_deleteReview + Sum_deleteWork + Sum_deleteBack;

            appendRow(new { value = "總計", style = "FooterTableCell", className = "CenterWord" },
                new { value = Sum_addReview, style = "FooterTableCell", className= "EndWord" },
                new { value = Sum_addWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_addBack, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateReview, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateBack, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteReview, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteBack, style = "FooterTableCell", className = "EndWord" },
                new { value = sum_total, style = "FooterTableCell", className = "EndWord" });//統計資料數

        }
    }
}

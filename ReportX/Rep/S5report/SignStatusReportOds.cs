using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.S5report
{
    public class SignStatusReportOds: SignStatusOds
    {
        string customCSS = @"  <office:automatic-styles>
    <style:style style:name='CenterWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='DataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0cm'/>
    </style:style>
    <style:style style:name='FirstDataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='none' fo:border-left='thin solid #000000' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='EndDataWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='none' fo:border-left='none' fo:border-right='thin solid #000000' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0cm'/>
    </style:style>
    <style:style style:name='TotalWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='none' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0cm'/>
    </style:style>
    <style:style style:name='TotalEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='none' fo:border-right='thin solid #000000' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='end' fo:margin-right='0cm'/>
    </style:style>
    <style:style style:name='ColumnFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='thin solid #000000' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='ColumnWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='none' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='ColumnEndWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='thin solid #000000' fo:border-bottom='none' fo:border-left='none' fo:border-right='thin solid #000000' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TotalFirstWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties fo:border-top='none' fo:border-bottom='thin solid #000000' fo:border-left='thin solid #000000' fo:border-right='none' style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
    </style:style>
    <style:style style:name='TitleWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='標楷體' style:font-name-asian='標楷體' style:font-name-complex='標楷體' fo:font-size='18pt' style:font-size-asian='18pt' style:font-size-complex='18pt' style:font-family-generic='script'/>
    </style:style>
    <style:style style:name='DataRangeWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='Times New Roman' style:font-name-asian='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='14pt' style:font-size-asian='14pt' style:font-size-complex='14pt'/>
    </style:style>
    <style:style style:name='UserWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' fo:wrap-option='wrap' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='center'/>
      <style:text-properties style:font-name='Times New Roman' style:font-name-asian='Times New Roman' style:font-name-complex='Times New Roman' fo:font-size='8pt' style:font-size-asian='8pt' style:font-size-complex='8pt'/>
    </style:style>
    <style:style style:name='ce17' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
      <style:table-cell-properties style:vertical-align='middle' style:repeat-content='false'/>
      <style:paragraph-properties fo:text-align='start' fo:margin-left='0cm'/>
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
        public SignStatusReportOds(Type model) : base(model)
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
                new { value = Sum_addReview, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_addWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_addBack, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateReview, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_updateBack, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteReview, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteWork, style = "FooterTableCell", className = "EndWord" },
                new { value = Sum_deleteBack, style = "FooterTableCell", className = "EndWord" },
                new { value = sum_total, style = "FooterEndTableCell", className = "EndWord" });//統計資料數

        }
    }
}

using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.OpenOffice;
using System;
using System.Data;
using System.Linq;

namespace ReportX
{
    /// <summary>
    /// 標準報表產生器
    /// </summary>
    /// <typeparam name="T">產出的報表類型</typeparam>
    public class ReportCreator<T> where T : IReportX, new()
    {
        /// <summary>
        /// 報表物件
        /// </summary>
        public T report { get; set; }

        /// <summary>
        /// 建構子，建立報表物件
        /// </summary>
        public ReportCreator()
        {
            report = new T();
        }

        /// <summary>
        /// 顯示報表資料時間範圍
        /// </summary>
        /// <param name="from">開始時間</param>
        /// <param name="to">結束時間，預設現在</param>
        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;
            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");
            if (report is AbsOffice)
                report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
            else if (report is AbsOpenOffice)
                report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), "TableCellData", "TitleDateWord");
        }

        /// <summary>
        /// 顯示報表建立人員
        /// </summary>
        /// <param name="creator">建立人員姓名</param>
        public void setCreator(string creator)
        {
            if (creator == null) return;
            report.setData(author: creator);
            if (report is AbsOffice)
                report.appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
            else if (report is AbsOpenOffice)
                report.appendFullRow(string.Format("製表人：{0}", creator), "TableCellData", "TitleTimeWord");
        }

        /// <summary>
        /// 設定報表名稱(如果是試算表會顯示在頁籤上)
        /// </summary>
        /// <param name="name">報表名稱</param>
        public void setSheetName(string name)
        {
            report.setData(sheetName: name);
        }


        /// <summary>
        /// 顯示報表標題
        /// </summary>
        /// <param name="title">標題文字</param>
        public void setTile(string title)
        {
            report.setData(sheetName: title);
            if (report is AbsOffice)
                report.appendFullRow(title, null, "r-header-title");
            else if (report is AbsOpenOffice)
                report.appendFullRow(title, "TableCellData", "Title");
        }

        /// <summary>
        /// 顯示報表建立時間 (現在)
        /// </summary>
        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            if (report is AbsOffice)
                report.appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
            else if (report is AbsOpenOffice)
                report.appendFullRow(string.Format("製表時間：{0}", now), "TableCellData", "TitleTimeWord");
        }

        /// <summary>
        /// 顯示報表欄位說明列，需要在 setData 之後才能呼叫
        /// </summary>
        public void setColumn()
        {
            ModelTR col = report.appendRow(report.cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }

        /// <summary>
        /// 顯示報表表格資料
        /// </summary>
        /// <param name="data">模型陣列式資料</param>
        public void setData<R>(R[] data)
        {
            report.appendTable(data);
        }
        
        /// <summary>
        /// 設定報表表格資料
        /// </summary>
        /// <param name="data">DataTable 格式的資料</param>
        public void setData(DataTable data)
        {
            report.appendTable(data);
        }

        /// <summary>
        /// 顯示報表表格資料總筆數 
        /// </summary>
        /// <param name="data">資料模型陣列</param>
        public void setSum<R>(R[] data)
        {
            _setSum(data.Length);
        }

        /// <summary>
        /// 顯示報表表格資料總筆數 
        /// </summary>
        /// <param name="data">DataTable 資料</param>
        public void setSum(DataTable data)
        {
            _setSum(data.Select().Count());
        }

        /// <summary>
        /// 集中優化，顯示報表表格資料總筆數 
        /// </summary>
        /// <param name="count">資料筆數</param>
        private void _setSum(int count)
        {
            if (report is AbsOffice)
            {
                report.appendRow(new
                {
                    value = "總筆數",
                    colspan = report.getColCount() - 1,
                    style = "background-color:#DDD;-webkit-print-color-adjust: exact;"
                }, count);
            }
            else if (report is AbsOpenOffice)
            {
                report.appendRow(new
                {
                    value = count,
                    colspan = report.getColCount() - 1,
                    style = "TotalCell",
                    className = "Word"
                });
            }
        }

        /// <summary>
        /// 需要在 setData 之後才能呼叫，過濾資料的顯示欄位，從既有的欄位定義中，過濾出指定顯示的欄位
        /// </summary>
        /// <param name="cut">欄位陣列</param>
        public void setFileterColumn(string[] cut)
        {
            report.changecut(cut);
        }


        /// <summary>
        /// 設定預設報表的CSS樣式
        /// </summary>
        public void setDefaultCss()
        {
            if (report is AbsOffice) report.setCustomStyle(customOfficeCSS);
            else if (report is Odt) report.setCustomStyle(customOdtCss);
            else if (report is Ods) report.setCustomStyle(customOdsCss);
        }



        /// <summary>
        /// 設定標準報表 (使用資料模型)
        /// </summary>
        /// <param name="data">資料</param>
        /// <param name="cols">欲顯示的欄位</param>
        /// <param name="title">標題</param>
        /// <param name="from">資料開始時間</param>
        /// <param name="to">資料結束時間</param>
        /// <param name="creator">報表建立人</param>
        /// <param name="showTotal">是否顯示資料總數</param>
        public void setInfo<R>(R[] data, string[] cols, string title, DateTime from, DateTime? to = null, string creator = null, bool showTotal = false)
        {
            report.setCol(data);
            if (cols != null && cols.Length > 0) setFileterColumn(cols);
            setDefaultCss();
            setTile(title);
            setDate(from, to);
            setCreator(creator);
            setCreatedDate();
            setColumn();
            setData(data);
            if (showTotal) setSum(data);
        }


        /// <summary>
        /// 設定標準報表 (使用DataTable)
        /// </summary>
        /// <param name="data">資料</param>
        /// <param name="cols">欲顯示的欄位</param>
        /// <param name="title">標題</param>
        /// <param name="from">資料開始時間</param>
        /// <param name="to">資料結束時間</param>
        /// <param name="creator">報表建立人</param>
        /// <param name="showTotal">是否顯示資料總數</param>
        public void setInfo(DataTable data, string[] cols, string title, DateTime from, DateTime? to = null, string creator = null, bool showTotal = false)
        {
            report.setCol(data);
            if (cols != null && cols.Length > 0) setFileterColumn(cols);
            setDefaultCss();
            setTile(title);
            setDate(from, to);
            setCreator(creator);
            setCreatedDate();
            setColumn();
            setData(data);
            if (showTotal) setSum(data);
        }

        /// <summary>
        /// 計算出報表字串
        /// </summary>
        /// <returns>報表結果字串</returns>
        public string render()
        {
            return report.render();
        }



        const string customOfficeCSS = @"
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
            }";
        const string customOdtCss = @"
            <office:automatic-styles>
            <style:style style:name='TableColumn' style:family='table-column'>
              <style:table-column-properties style:column-width='auto'/>
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
          </office:master-styles>";
        const string customOdsCss = @"<office:automatic-styles>
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
          </office:master-styles>";
    }
}
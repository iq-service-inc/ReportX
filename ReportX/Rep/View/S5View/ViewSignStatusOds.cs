using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{

        public class ViewSignStatusOds
        {
            private ModelSignStatusOds m;

            public ViewSignStatusOds(ModelSignStatusOds model)
            {
                m = model;
            }
            public string render()
            {
                string style = m.style.render(),
                       body = m.body.render();

                // more coustom code here
                // ...

                return string.Format(template, m.author, m.company, m.sheetName, m.datetime, style, body, m.colNum, m.dateRange);

            }
        string template =
 @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<office:document-content xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' xmlns:xlink='http://www.w3.org/1999/xlink' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:number='urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0' xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0' xmlns:of='urn:oasis:names:tc:opendocument:xmlns:of:1.2' office:version='1.2'>
{4} 
 <office:body>
    <office:spreadsheet>
      <table:calculation-settings table:case-sensitive='false' table:search-criteria-must-apply-to-whole-cell='true' table:use-wildcards='true' table:use-regular-expressions='false' table:automatic-find-labels='false'/>
      <table:table table:name='{2}' table:style-name='ta1'>
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column'  />
        <table:table-column table:style-name='Column'  />
        <table:table-column table:style-name='Column'  />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column'  />
        <table:table-row table:style-name='Row'>
          <table:table-cell office:value-type='string' table:number-columns-spanned='12' table:number-rows-spanned='1' table:style-name='TitleWord'>
            <text:p>{2}</text:p>
          </table:table-cell>
          <table:covered-table-cell />
          <table:table-cell />
        </table:table-row>
        <table:table-row table:style-name='Row'>
          <table:table-cell office:value-type='string' table:number-columns-spanned='12' table:number-rows-spanned='1' table:style-name='DataRangeWord'>
            <text:p>{7}</text:p>
          </table:table-cell>
          <table:covered-table-cell />
          <table:table-cell />
        </table:table-row>
        <table:table-row table:style-name='Row'>
          <table:table-cell office:value-type='string' table:number-columns-spanned='12' table:number-rows-spanned='1' table:style-name='UserWord'>
            <text:p>製表人：{0}</text:p>
          </table:table-cell>
          <table:covered-table-cell />
          <table:table-cell />
        </table:table-row>
        <table:table-row table:style-name='Row'>
          <table:table-cell office:value-type='string' table:number-columns-spanned='12' table:number-rows-spanned='1' table:style-name='UserWord'>
            <text:p>製表時間：{3}</text:p>
          </table:table-cell>
          <table:covered-table-cell />
          <table:table-cell />
        </table:table-row>
         {5}
 <table:table-row table:style-name='Row'>
          <table:table-cell office:value-type='string' table:style-name='ce17'>
            <text:p>※僅顯示有送審知識的日期資料。</text:p>
          </table:table-cell>
          <table:table-cell  table:style-name='ce1'/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>";
        
    }
}

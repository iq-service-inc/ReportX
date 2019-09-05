using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewSignStatus
    {
        private ModelSignStatus m;

        public ViewSignStatus(ModelSignStatus model)
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
<office:document-content xmlns:anim='urn:oasis:names:tc:opendocument:xmlns:animation:1.0' xmlns:chart='urn:oasis:names:tc:opendocument:xmlns:chart:1.0' xmlns:config='urn:oasis:names:tc:opendocument:xmlns:config:1.0' xmlns:db='urn:oasis:names:tc:opendocument:xmlns:database:1.0' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dr3d='urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' xmlns:form='urn:oasis:names:tc:opendocument:xmlns:form:1.0' xmlns:grddl='http://www.w3.org/2003/g/data-view#' xmlns:math='http://www.w3.org/1998/Math/MathML' xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0' xmlns:number='urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:presentation='urn:oasis:names:tc:opendocument:xmlns:presentation:1.0' xmlns:script='urn:oasis:names:tc:opendocument:xmlns:script:1.0' xmlns:smil='urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0' xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0' xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:xforms='http://www.w3.org/2002/xforms' xmlns:xhtml='http://www.w3.org/1999/xhtml' xmlns:xlink='http://www.w3.org/1999/xlink' office:version='1.2'> 
{4} 
<office:body>
    <office:text text:use-soft-page-breaks='true'>
      <text:p text:style-name='Title'>{2}</text:p>
      <text:p text:style-name='DateRange'>{7} </text:p>
      <text:p text:style-name='User'>
        製表人：{0}<text:line-break/>製表時間：{3}
      </text:p>
      <table:table table:style-name='Table'>
        <table:table-columns>
          <table:table-column table:style-name='TableColumn5'/>
          <table:table-column table:style-name='TableColumn6'/>
          <table:table-column table:style-name='TableColumn7'/>
          <table:table-column table:style-name='TableColumn8'/>
          <table:table-column table:style-name='TableColumn9'/>
          <table:table-column table:style-name='TableColumn10'/>
          <table:table-column table:style-name='TableColumn11'/>
          <table:table-column table:style-name='TableColumn12'/>
          <table:table-column table:style-name='TableColumn13'/>
          <table:table-column table:style-name='TableColumn14'/>
          <table:table-column table:style-name='TableColumn15'/>
          <table:table-column table:style-name='TableColumn16'/>
        </table:table-columns>
         {5}
      </table:table>
      <table:table table:style-name='NoticeTable'>
        <table:table-columns>
          <table:table-column table:style-name='TableColumn102'/>
        </table:table-columns>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='NoticeTableRow'>
            <text:p text:style-name='Word'>※僅顯示有送審知識的日期資料。</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
      <text:p text:style-name='Word'/>
    </office:text>
  </office:body>
</office:document-content>";
    }
}

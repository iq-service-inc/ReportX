using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewAmountOds
    {
        private ModelAmountOds m;

        public ViewAmountOds(ModelAmountOds model)
        {
            m = model;
        }
        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(template, m.author, m.company, m.sheetName, m.datetime, style, body);

        }
        string template =
 @"<?xml version='1.0' encoding='UTF-8'?>
<office:document-content xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' xmlns:xlink='http://www.w3.org/1999/xlink' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0' xmlns:number='urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0' xmlns:presentation='urn:oasis:names:tc:opendocument:xmlns:presentation:1.0' xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0' xmlns:chart='urn:oasis:names:tc:opendocument:xmlns:chart:1.0' xmlns:dr3d='urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0' xmlns:math='http://www.w3.org/1998/Math/MathML' xmlns:form='urn:oasis:names:tc:opendocument:xmlns:form:1.0' xmlns:script='urn:oasis:names:tc:opendocument:xmlns:script:1.0' xmlns:ooo='http://openoffice.org/2004/office' xmlns:ooow='http://openoffice.org/2004/writer' xmlns:oooc='http://openoffice.org/2004/calc' xmlns:dom='http://www.w3.org/2001/xml-events' xmlns:xforms='http://www.w3.org/2002/xforms' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:rpt='http://openoffice.org/2005/report' xmlns:of='urn:oasis:names:tc:opendocument:xmlns:of:1.2' xmlns:xhtml='http://www.w3.org/1999/xhtml' xmlns:grddl='http://www.w3.org/2003/g/data-view#' xmlns:tableooo='http://openoffice.org/2009/table' xmlns:field='urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0' office:version='1.2'>
 {4}
  <office:body>
    <office:spreadsheet>
      <table:calculation-settings table:case-sensitive='false' table:automatic-find-labels='false' table:use-regular-expressions='false'/>
      <table:table table:name='{2}' table:style-name='ta1'>
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-column table:style-name='Column' />
        <table:table-row table:style-name='Row'>
          <table:table-cell table:style-name='TableName' office:value-type='string' table:number-columns-spanned='4' table:number-rows-spanned='1'>
            <text:p>{2}</text:p>
          </table:table-cell>
          <table:covered-table-cell table:number-columns-repeated='3'/>
          <table:table-cell />
        </table:table-row>
        <table:table-row table:style-name='Row'>
          <table:table-cell table:style-name='User' office:value-type='string' table:number-columns-spanned='4' table:number-rows-spanned='1'>
            <text:p>製表人：{0}</text:p>
          </table:table-cell>
          <table:covered-table-cell table:number-columns-repeated='3'/>
          <table:table-cell />
        </table:table-row>
        <table:table-row table:style-name='Row'>
          <table:table-cell table:style-name='User' office:value-type='string' table:number-columns-spanned='4' table:number-rows-spanned='1'>
            <text:p>製表時間：{3}</text:p>
          </table:table-cell>
          <table:covered-table-cell table:number-columns-repeated='3'/>
          <table:table-cell />
        </table:table-row>
     {5}
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>";
    }
}

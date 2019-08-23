using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewOdt
    {
        private ModelOdt m;

        public ViewOdt(ModelOdt model)
        {
            m = model;
        }

        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(wordtest,m.author, m.company, m.sheetName, style, body);

        }
        string wordtest = @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<office:document-content xmlns:anim='urn:oasis:names:tc:opendocument:xmlns:animation:1.0' xmlns:chart='urn:oasis:names:tc:opendocument:xmlns:chart:1.0' xmlns:config='urn:oasis:names:tc:opendocument:xmlns:config:1.0' xmlns:db='urn:oasis:names:tc:opendocument:xmlns:database:1.0' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dr3d='urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' xmlns:form='urn:oasis:names:tc:opendocument:xmlns:form:1.0' xmlns:grddl='http://www.w3.org/2003/g/data-view#' xmlns:math='http://www.w3.org/1998/Math/MathML' xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0' xmlns:number='urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:presentation='urn:oasis:names:tc:opendocument:xmlns:presentation:1.0' xmlns:script='urn:oasis:names:tc:opendocument:xmlns:script:1.0' xmlns:smil='urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0' xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0' xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:xforms='http://www.w3.org/2002/xforms' xmlns:xhtml='http://www.w3.org/1999/xhtml' xmlns:xlink='http://www.w3.org/1999/xlink' office:version='1.2'>
{3}
  <office:body>
    <office:text text:use-soft-page-breaks='true'>
      <table:table table:style-name='Table'>
        <table:table-columns>
          <table:table-column table:style-name='TableColumn'/>
          <table:table-column table:style-name='TableColumn'/>
          <table:table-column table:style-name='TableColumn'/>
          <table:table-column table:style-name='TableColumn'/>
          <table:table-column table:style-name='TableColumn'/>
          <table:table-column table:style-name='TableColumn'/>
        </table:table-columns>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='TableCellData' table:number-columns-spanned='6'>
            <text:p text:style-name='Title'>
              <text:span text:style-name='TitleWord'>標題</text:span>
            </text:p>
          </table:table-cell>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
        </table:table-row>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='TableCellData' table:number-columns-spanned='6'>
            <text:p text:style-name='TitleDateWord'>2019/08/21 - 2019/08/22</text:p>
          </table:table-cell>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
        </table:table-row>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='TableCellData' table:number-columns-spanned='6'>
            <text:p text:style-name='TitleTimeWord'>製表時間：2019/08/22 02:45:下午</text:p>
          </table:table-cell>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
        </table:table-row>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>ID</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>標題</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>姓名</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>編號</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>資料</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='TitleCell'>
            <text:p text:style-name='TitleCellWord'>電話</text:p>
          </table:table-cell>
        </table:table-row>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>100</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>測試_0</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>SOL_0</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>
              123<text:s/>
            </text:p>
          </table:table-cell>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>data0</text:p>
          </table:table-cell>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>09234567890</text:p>
          </table:table-cell>
        </table:table-row>
        <table:table-row table:style-name='TableRow'>
          <table:table-cell table:style-name='TotalCell' table:number-columns-spanned='5'>
            <text:p text:style-name='Word'>總筆數</text:p>
          </table:table-cell>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:covered-table-cell/>
          <table:table-cell table:style-name='CellWord'>
            <text:p text:style-name='Word'>1</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
    </office:text>
  </office:body>
</office:document-content>
";
    }
}

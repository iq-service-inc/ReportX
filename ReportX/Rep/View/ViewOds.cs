using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewOds
    {
        private ModelOds m;

        public ViewOds(ModelOds model)
        {
            m = model;
        }

        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(wordtest, m.author, m.company, m.sheetName, style, body);

        }
        string wordtest = @"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<office:document-content xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' xmlns:xlink='http://www.w3.org/1999/xlink' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:number='urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0' xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0' xmlns:of='urn:oasis:names:tc:opendocument:xmlns:of:1.2' office:version='1.2'>
{3}
  <office:body>
    <office:spreadsheet>
      <table:calculation-settings table:case-sensitive='false' table:search-criteria-must-apply-to-whole-cell='true' table:use-wildcards='true' table:use-regular-expressions='false' table:automatic-find-labels='false'/>
      <table:table table:name='{2}' table:style-name='ta1'>
       {4}
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
";
    }
}

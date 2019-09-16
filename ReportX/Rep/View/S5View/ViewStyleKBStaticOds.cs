using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewStyleKBStaticOds
    {
        private string costomCSS = "";
        private string patchCSS = "";

        public ViewStyleKBStaticOds()
        {
            patchCSS = @"  <office:font-face-decls>
    <style:font-face style:name='新細明體' svg:font-family='新細明體'/>
    <style:font-face style:name='標楷體' svg:font-family='標楷體'/>
    <style:font-face style:name='Times New Roman' svg:font-family='&quot;Times New Roman&quot;'/>
  </office:font-face-decls>";
        }

        public void setCustomCSS(string costomCSS)
        {
            this.costomCSS = costomCSS;
        }




        public string render()
        {
            string format_all_css = string.Format(global_css, patchCSS, costomCSS);
            return string.Format(template, format_all_css);
        }

        string template = @"{0}";

        string global_css = @"
            {0}
            {1}
        ";
    }
}

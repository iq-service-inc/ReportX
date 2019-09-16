using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewStyleAmountOds
    {
        private string costomCSS = "";
        private string patchCSS = "";

        public ViewStyleAmountOds()
        {
            patchCSS = @"  <office:font-face-decls>
    <style:font-face style:name='Times New Roman' svg:font-family='&apos;Times New Roman&apos;'/>
    <style:font-face style:name='新細明體' svg:font-family='新細明體'/>
    <style:font-face style:name='標楷體1' svg:font-family='標楷體'/>
    <style:font-face style:name='標楷體' svg:font-family='標楷體' style:font-family-generic='script'/>
    <style:font-face style:name='Arial' svg:font-family='Arial' style:font-family-generic='swiss' style:font-pitch='variable'/>
    <style:font-face style:name='Lucida Sans Unicode' svg:font-family='&apos;Lucida Sans Unicode&apos;' style:font-family-generic='system' style:font-pitch='variable'/>
    <style:font-face style:name='Tahoma' svg:font-family='Tahoma' style:font-family-generic='system' style:font-pitch='variable'/>
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

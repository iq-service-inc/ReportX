using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportX.Rep.Model;

namespace ReportX.Rep.View
{
    public class ViewStyleOdt
    {
        private string costomCSS = "";
        private string patchCSS = "";

        public ViewStyleOdt()
        {
            patchCSS = @"<office:font-face-decls>
    <style:font-face style:name='Calibri' svg:font-family='Calibri' style:font-family-generic='swiss' style:font-pitch='variable' svg:panose-1='2 15 5 2 2 2 4 3 2 4'/>
    <style:font-face style:name='新細明體' svg:font-family='新細明體' style:font-family-generic='roman' style:font-pitch='variable' svg:panose-1='2 2 5 0 0 0 0 0 0 0'/>
    <style:font-face style:name='Times New Roman' svg:font-family='Times New Roman' style:font-family-generic='roman' style:font-pitch='variable' svg:panose-1='2 2 6 3 5 4 5 2 3 4'/>
    <style:font-face style:name='Tahoma' svg:font-family='Tahoma' style:font-family-generic='swiss' style:font-pitch='variable' svg:panose-1='2 11 6 4 3 5 4 4 2 4'/>
    <style:font-face style:name='微軟正黑體' svg:font-family='微軟正黑體' style:font-family-generic='swiss' style:font-pitch='variable' svg:panose-1='2 11 6 4 3 5 4 4 2 4'/>
    <style:font-face style:name='Calibri Light' svg:font-family='Calibri Light' style:font-family-generic='swiss' style:font-pitch='variable' svg:panose-1='2 15 3 2 2 2 4 3 2 4'/>
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

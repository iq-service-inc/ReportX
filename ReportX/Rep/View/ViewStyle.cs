using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewStyle
    {
        private string costomCSS = "";
        private string patchCSS = "";

        public ViewStyle()
        {
            renderPatchCSS();
        }

        public void setCustomCSS(string costomCSS)
        {
            this.costomCSS = costomCSS;
        }

        private void renderPatchCSS()
        {
            for (int i = 8; i <= 36; i++)
                patchCSS += string.Format(template_fontSizeClass, i);
            // more useful css here
        }


        public string render()
        {
            string format_all_css = string.Format(global_css, patchCSS, costomCSS);
            return string.Format(template, format_all_css);
        }

        string template = @"
            <style>
                <!--        
                {0}
                -->
                @media print{{
                {0}
                }} 
            </style>
        ";

        string global_css = @"
            table{{
                mso-number-format:'\@';
                mso-displayed-decimal-separator:'\.';
                mso-displayed-thousand-separator:'\,';
                font-family: 微軟正黑體, Microsoft JhengHei, PMingLiU, serif;
                border-collapse: collapse;
                table-layout:fixed;
            }}
            @page{{
                margin:1.0in .75in 1.0in .75in;
                mso-header-margin:.5in;
                mso-footer-margin:.5in;
            }}
            tr{{
                mso-height-source:auto;
                mso-ruby-visibility:none; 
            }}
            table, th, td{{
                border: 1px #AAA solid;
                padding: 0px 2px;
                white-space: pre;
            }}
            col{{
                mso-width-source:auto;
                mso-ruby-visibility:none;
            }}
            br{{
                mso-data-placement:same-cell;
            }}
            .tac{{
                text-align: center;
            }}
            {0}
            {1}
        ";

        string template_fontSizeClass = ".fz{0}{{font-size:{0}px;}}";
    }
}

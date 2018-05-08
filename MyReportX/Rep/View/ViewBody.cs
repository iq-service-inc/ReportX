using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyReportX.Rep.Model;

namespace MyReportX.Rep.View
{
    public class ViewBody
    {
        private List<ModelTR> model;
        private int? width; // if not set, keep it is auto

        public ViewBody(List<ModelTR> model, int? width = null)
        {
            this.model = model;
            this.width = width;
        }

        public string render()
        {
            string table_width = width == null ? "" : string.Format("width={0}", width),
                   trs = "";

            foreach (ModelTR tr in model)
            {
                string tr_className = tr.className == null ? "" : string.Format("class=\"{0}\" ", tr.className),
                       tr_customStyle = tr.style ?? "",
                       tr_style = string.Format("style=\"{0}\" ", tr_customStyle) + tr_className,
                       tds = "";

                if (tr.tds != null)
                {
                    foreach (ModelTD td in tr.tds)
                    {
                        if (td == null) continue;
                        string attributes = "",
                               td_style = "",
                               className = td.className == null ? "" : td.className,
                               data = td.data == null ? "" : td.data.ToString();

                        if (td.colspan > 1)
                            attributes += string.Format("colspan={0} ", td.colspan);
                        if (td.rowspan > 1)
                            attributes += string.Format("rowspan={0} ", td.rowspan);

                        if (td.fontSize != null)
                            td_style += string.Format("font-size:{0};", td.fontSize);

                        if (td.align != null)
                            td_style += string.Format("text-align:{0};", td.align);

                        if (td.bold)
                            td_style += string.Format("font-weight:{0};", td.bold);

                        if (td.style != null)
                            td_style += td.style;

                        attributes += string.Format("style=\"{0}\" ", td_style);
                        attributes += string.Format("class=\"{0}\" ", className);

                        tds += string.Format(template_td, attributes, data);
                    }
                }
                trs += string.Format(template_tr, tr_style, tds);
            }

            return string.Format(template, trs, table_width);
        }

        string template = @"
            <div align=center x:publishsource='Excel'>
                <table x:str {1}>
                {0}
                </table>
            </div> 
        ";

        string template_td = "<td {0}>{1}</td>";
        string template_tr = "<tr {0}>{1}</tr>";

    }
}

using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewBodyKBStatic
    {
        private List<ModelTR> model;
        public MemberInfo[] modeli;
        private int? width; // if not set, keep it is auto
        public ViewBodyKBStatic(List<ModelTR> model, int? width = null)
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
                       tds = ""
                      ;

                if (tr.tds != null)
                {
                    foreach (ModelTD td in tr.tds)
                    {
                        if (td == null) continue;
                        string attributes = "",
                               text_style = "",
                               td_style = "",
                               table_cell = "",
                               className = td.className == null ? "" : td.className,
                               data = td.data == null ? "" : td.data.ToString();

                        if (td.className == "column")
                        {
                            attributes += string.Format("table:style-name='TableCell'");
                            className = "Data";
                        }
                        if (td.className == null)
                        {
                            attributes += string.Format("table:style-name='DataTableCell'");
                            switch (td.col)
                            {
                                case "編號":
                                    className = "CenterWord";
                                    break;
                                case "知識目錄":
                                    className = "Data";
                                    break;
                                case "知識標題":
                                    className = "Data";
                                    break;
                                case "建立時間":
                                    className = "CenterWord";
                                    break;
                                case "建立人員":
                                    className = "CenterWord";
                                    break;
                                default:
                                    className = "Data";
                                    break;
                            }
                        }
                        if (td_style == null)
                            attributes += string.Format("table:style-name=\"{0}\" ", td_style);
                        text_style += string.Format("text:style-name=\"{0}\" ", className);
                        tds += string.Format(template_td, attributes, text_style, data, table_cell);
                    }
                }
                trs += string.Format(template_tr, tr_style, tds);
            }

            return string.Format(template, trs, table_width);
        }


        string template = @"{0}";
        string template_td = "<table:table-cell  {0}><text:p {1}>{2}</text:p></table:table-cell>{3} ";
        string template_tr = "<table:table-row table:style-name='TableRow '>{1}</table:table-row>";

    }
}

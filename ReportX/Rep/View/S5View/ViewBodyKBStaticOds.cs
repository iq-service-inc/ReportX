using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewBodyKBStaticOds
    {
        private List<ModelTR> model;
        public MemberInfo[] modeli;
        private int? width; // if not set, keep it is auto
        public ViewBodyKBStaticOds(List<ModelTR> model, int? width = null)
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
                            switch (td.data)
                            {
                                case "編號":
                                    attributes += string.Format(" office:value-type='string' table:style-name='ColumnFirstWord'");
                                    break;
                                case "建立人員":
                                    attributes += string.Format("office:value-type='string' table:style-name='ColumnEndWord'");
                                    break;
                                default:
                                    attributes += string.Format("office:value-type='string' table:style-name='ColumnCenterWord'");
                                    break;
                            }
                        }
                        if (td.className == null)
                        {
                            switch (td.col)
                            {
                                case "編號":
                                    attributes += string.Format(" office:value-type='float' office:value='{0}' table:style-name='FirstDataWord'",td.data);
                                    break;
                                case "知識目錄":
                                    attributes += string.Format(" office:value-type='string' office:value='{0}' table:style-name='FirstDataWord'", td.data);
                                    break;
                                case "知識標題":
                                    attributes += string.Format(" office:value-type='string' office:value='{0}' table:style-name='FirstDataWord'", td.data);
                                    break;
                                case "建立時間":
                                    attributes += string.Format(" office:value-type='string' office:value='{0}' table:style-name='DataCenterWord'", td.data);
                                    break;
                                case "建立人員":
                                    attributes += string.Format(" office:value-type='string' office:value='{0}' table:style-name='EndDataWord'", td.data);
                                    break;
                                default:
                                    className = "Data";
                                    break;
                            }
                        }
                        if (td_style == null)
                            attributes += string.Format("table:style-name=\"{0}\" ", td_style);
                        tds += string.Format(template_td, attributes, data, table_cell);
                    }
                }
                trs += string.Format(template_tr, tr_style, tds);
            }

            return string.Format(template, trs, table_width);
        }


        string template = @"{0}";
        string template_td = "<table:table-cell  {0}><text:p>{1}</text:p></table:table-cell>{2} ";
        string template_tr = "<table:table-row table:style-name='Row '>{1}</table:table-row>";
    }
}

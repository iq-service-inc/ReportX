using ReportX.Rep.Attributes;
using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewBodyAmount
    {
        private List<ModelTR> model;
        public MemberInfo[] modeli;
        private int? width; // if not set, keep it is auto
        public ViewBodyAmount(List<ModelTR> model, int? width = null)
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
                       tds = "",
                       table_header =""
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
                            table_header = "</table:table-header-rows>";
                        }
                        if (td.className == null)
                        {
                            switch (td.col)
                            {
                                case "順序":
                                    attributes += string.Format("table:style-name='SequenceDataTableCell'");
                                    className = "Data";
                                    break;
                                case "知識目錄":
                                    attributes += string.Format("table:style-name='ContentDataTableCell'");
                                    className = "ContentData";
                                    var test = td.data.ToString().Substring(0, 1);
                                    if (test == "◎")
                                    {
                                        className = "MarkContentData";
                                    }
                                    else
                                    {
                                        className = "ContentData";
                                    }
                                    break;
                                case "有效知識":
                                    attributes += string.Format("table:style-name='KnowledgeDataTableCell'");
                                    className = "KnowledgeData";
                                    break;
                                case "無效知識":
                                    attributes += string.Format("table:style-name='KnowledgeDataTableCell'");
                                    className = "KnowledgeData";
                                    break;
                                default:
                                    attributes += string.Format("table:style-name='SequenceDataTableCell'");
                                    className = "Data";
                                    break;
                            }
                        }
                        if (td.style == "TotalCell")
                        {
                            table_cell = "";
                            attributes = "";
                            var lastCellStyle = "table:style-name='ContentTableCell'";
                            var lastTextStyle = "text:style-name='KnowledgeData'";
                            attributes += string.Format("table:style-name='ContentTableCell' table:number-columns-spanned=\"{0}\" ", td.colspan);
                            for (int i = 1; i < td.colspan; i++)
                            {

                                table_cell += "<table:covered-table-cell/>";
                            }
                            table_cell += string.Format(lastRow, lastCellStyle, lastTextStyle, td.sum_c);
                            table_cell += string.Format(lastRow, lastCellStyle, lastTextStyle, td.sum_w);
                        }
                        if (td_style == null)
                            attributes += string.Format("table:style-name=\"{0}\" ", td_style);
                        text_style += string.Format("text:style-name=\"{0}\" ", className);
                        tds += string.Format(template_td, attributes, text_style, data, table_cell);
                    }
                }
                trs += string.Format(template_tr, tr_style, tds, table_header);
            }

            return string.Format(template, trs, table_width);
        }


        string template = @"{0}";
        string template_td = "<table:table-cell  {0}><text:p {1}>{2}</text:p></table:table-cell>{3} ";
        string template_tr = "<table:table-row table:style-name='TitleTableRow '>{1}</table:table-row>{2}";
        string lastRow = "<table:table-cell  {0}><text:p {1}>{2}</text:p></table:table-cell>";
        //0:style-name='TableCellData' table:number-columns-spanned=\"{0}\"  1:text:style-name="Title" 2:text:style-name="TitleWord" 3:data
    }
}

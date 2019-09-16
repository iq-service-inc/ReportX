using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewBodyAmountOds
    {
        private List<ModelTR> model;
        public MemberInfo[] modeli;
        private int? width; // if not set, keep it is auto
        public ViewBodyAmountOds(List<ModelTR> model, int? width = null)
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
                                case "順序":
                                    attributes = string.Format("table:style-name='ColumneFirstWord' office:value-type='string'");
                                    break;
                                case "知識目錄":
                                    attributes = string.Format("table:style-name='ColumnWord' office:value-type='string'");
                                    break;
                                case "有效知識":
                                    attributes = string.Format("table:style-name='ColumnWord' office:value-type='string'");
                                    break;
                                case "無效知識":
                                    attributes = string.Format("table:style-name='ColumnEndWord' office:value-type='string'");
                                    break;
                                default:
                                    break;
                            }
                        }
                        if (td.className == null)
                        {
                            switch (td.col)
                            {
                                case "順序":
                                    attributes = string.Format("table:style-name='DataFirstWord' office:value-type='float' office:value='{0}'", td.data);
                                    break;
                                case "知識目錄":
                                    var test = td.data.ToString().Substring(0, 1);
                                    if (test == "◎")
                                    {
                                        attributes = string.Format("table:style-name='DataMarkWord' office:value-type='string'");
                                    }
                                    else
                                    {
                                        attributes = string.Format("table:style-name='DataWord' office:value-type='string'");
                                    }
                                    break;
                                case "有效知識":
                                    attributes = string.Format(" table:style-name='CorrectWord' office:value-type='float' office:value='{0}'", td.data);
                                    break;
                                case "無效知識":
                                    attributes = string.Format("table:style-name='DataEndWord' office:value-type='float' office:value='{0}'", td.data);
                                    break;
                                default:
                                    break;
                            }
                        }
                        if (td.className == "Total")
                        {
                            if (td.data == "合計")
                            {
                                attributes += string.Format("table:style-name='TotalFirstWord' office:value-type='string' table:number-columns-spanned='2' table:number-rows-spanned='1' ");
                                for (int i = 1; i < td.colspan; i++)
                                {
                                    table_cell += "<table:covered-table-cell/>";
                                }
                            }
                            else
                            {
                                switch (td.style)
                                {
                                    case "correct":
                                        attributes = string.Format("table:style-name='TotalCorrectWord' office:value-type='float' office:value='{0}'",td.data);
                                        break;
                                    case "wrong":
                                        attributes = string.Format("table:style-name='TotalWrongWord' office:value-type='float' office:value='{0}'",td.data);
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        tds += string.Format(template_td, attributes, data, table_cell);
                    }
                }
                trs += string.Format(template_tr, tr_style, tds);
            }

            return string.Format(template, trs, table_width);
        }


        string template = @"{0}";
        string template_td = "<table:table-cell  {0}><text:p >{1}</text:p></table:table-cell>{2} ";
        string template_tr = "<table:table-row table:style-name='Row'>{1}</table:table-row>";
    }
}

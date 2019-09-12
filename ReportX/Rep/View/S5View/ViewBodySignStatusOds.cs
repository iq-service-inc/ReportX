using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View.S5View
{
    public class ViewBodySignStatusOds
    {
        private List<ModelTR> model;
        public MemberInfo[] modeli;
        private int? width; // if not set, keep it is auto
        public ViewBodySignStatusOds(List<ModelTR> model, int? width = null)
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
                               td_style = td.style == null ? null : td.style,
                               table_cell = "",
                               className = td.className == null ? "" : td.className,
                               data = td.data == null ? "" : td.data.ToString();

                        if (td.className == "column")
                        {
                            switch (td.data)
                            {
                                case "順序":
                                    attributes += string.Format("office:value-type='string' table:number-columns-spanned='1' table:number-rows-spanned='2' table:style-name='ColumnFirstWord'");
                                    break;
                                case "審核起始日":
                                    attributes += string.Format("office:value-type='string' table:number-columns-spanned='1' table:number-rows-spanned='2' table:style-name='ColumnWord'");
                                    break;
                                case "新增知識":
                                    attributes += string.Format("office:value-type='string' table:number-columns-spanned='3' table:number-rows-spanned='1' table:style-name='ColumnWord'");
                                    table_cell += "<table:covered-table-cell/>";
                                    table_cell += "<table:covered-table-cell/>";
                                    break;
                                case "修改知識":
                                    attributes += string.Format("office:value-type='string' table:number-columns-spanned='3' table:number-rows-spanned='1' table:style-name='ColumnWord'");
                                    table_cell += "<table:covered-table-cell/>";
                                    table_cell += "<table:covered-table-cell/>";
                                    break;
                                case "刪除知識":
                                    attributes += string.Format(" office:value-type='string' table:number-columns-spanned='3' table:number-rows-spanned='1' table:style-name='ColumnWord'");
                                    table_cell += "<table:covered-table-cell/>";
                                    table_cell += "<table:covered-table-cell/>";
                                    break;
                                case "合計":
                                    attributes += string.Format("office:value-type='string' table:number-columns-spanned='1' table:number-rows-spanned='2' table:style-name='ColumnEndWord'");
                                    break;

                                default:
                                    continue;

                            }

                        }
                        if (td.className == "secondColumn")
                        {
                            switch (td.data)
                            {
                                case "順序":
                                    tds += " <table:covered-table-cell/>";
                                    continue;
                                case "審核起始日":
                                    tds += " <table:covered-table-cell/>";
                                    continue;

                                case "合計":
                                    tds += " <table:covered-table-cell/>";
                                    continue;
                                case "新增知識審核":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "審核中";
                                    break;
                                case "新增知識生效":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "生效";
                                    break;

                                case "新增知識退件":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "退件";
                                    break;
                                case "修改知識審核":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "審核中";
                                    break;
                                case "修改知識生效":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "生效";
                                    break;

                                case "修改知識退件":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "退件";
                                    break;
                                case "刪除知識審核":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "審核中";
                                    break;
                                case "刪除知識生效":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "生效";
                                    break;

                                case "刪除知識退件":
                                    attributes += string.Format(" office:value-type='string' table:style-name='CenterWord'");
                                    data = "退件";
                                    break;
                                default:
                                    continue;

                            }

                        }
                        if (td.className == null)
                        {
                            if (td.data != null)
                            {
                                switch (td.col)
                                {
                                    case "順序":
                                        attributes += string.Format("office:value-type='float' office:value='{0}' table:style-name='FirstDataWord'",td.data);
                                        break;
                                    case "審核起始日":
                                        attributes += string.Format("office:value-type='string' table:style-name='CenterWord'");
                                        break;
                                    case "合計":
                                        attributes += string.Format("office:value-type='float' office:value='{0}' table:style-name='EndDataWord'", td.data);
                                        break;
                                    default:
                                        attributes += string.Format("office:value-type='float' office:value='{0}' table:style-name='DataWord'", td.data);
                                        break;

                                }

                            }
                            else
                            {
                                continue;
                            }
                        }
                        if (td_style != null)
                            attributes += string.Format("table:style-name=\"{0}\" ", td_style);
                        if (td_style == "FooterTableCell")
                            attributes = string.Format("office:value-type='float' office:value='{0}' table:style-name='TotalWord'", td.data);
                        if (td_style == "FooterEndTableCell")
                            attributes = string.Format("office:value-type='float' office:value='{0}' table:style-name='TotalEndWord'", td.data);
                        if (td.data == "總計")
                        {
                            attributes = " office:value-type='string' table:number-columns-spanned='2' table:number-rows-spanned='1' table:style-name='TotalFirstWord'";
                            table_cell += "<table:covered-table-cell/>";
                        }
                        text_style += string.Format("text:style-name=\"{0}\" ", className);
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

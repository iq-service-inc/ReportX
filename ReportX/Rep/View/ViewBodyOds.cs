using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewBodyOds
    {
        private List<ModelTR> model;
        private int? width; // if not set, keep it is auto
        public ViewBodyOds(List<ModelTR> model, int? width = null)
        {
            this.model = model;
            this.width = width;
        }
        public string render()
        {
            string table_width = width == null ? "" : string.Format("width={0}", width),
                   trs = "";
            if (width != null)
            {
                table_width = "";
                for (int i = 0; i < width; i++)
                {

                    table_width += "<table:table-column table:style-name='TableColumn' />";
                }
            }
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
                               text_style = "",
                               td_style = "",
                               table_cell = "",
                               className = td.className == null ? "" : td.className,
                               data = td.data == null ? "" : td.data.ToString();

                        if (td.colspan > 1)
                        {
                            attributes += string.Format("office:value-type='string' table:style-name='TitleWord'  table:number-rows-spanned='1' table:number-columns-spanned=\"{0}\" ", td.colspan);
                                table_cell = "<table:covered-table-cell/>";
                            
                        }
                        if (td.className == "column")
                        {
                            attributes += string.Format("table:style-name='ColumnWord' office:value-type='string'");
                        }
                        if (td.className == null)
                        {
                            attributes += string.Format("table:style-name='Word' office:value-type='string'");
                        }
                        if (td.style == "TotalCell")
                        {
                            table_cell = "";
                            attributes = "";
                            var lastCellStyle = "office:value-type='string' table:style-name='Word'";
                            var lastData = data;
                            attributes += string.Format("office:value-type='string' table:number-columns-spanned=\"{0}\" table:number-rows-spanned='1' table:style-name='TotalWord'", td.colspan);
                            for (int i = 1; i < td.colspan; i++)
                            {

                                table_cell += "<table:covered-table-cell/>";
                            }
                            table_cell += string.Format(lastRow, lastCellStyle, lastData);
                            data = "總筆數";
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


        string template = @" {1}{0}";

        string template_td = "<table:table-cell  {0}><text:p>{1}</text:p></table:table-cell> {2}";
        string template_tr = "<table:table-row table:style-name='TableRow'>{1}</table:table-row>";
        string lastRow = "<table:table-cell  {0}><text:p>{1}</text:p></table:table-cell>";

    }
}

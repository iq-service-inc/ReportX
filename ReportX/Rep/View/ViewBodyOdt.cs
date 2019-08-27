using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReportX.Rep.View
{
    public class ViewBodyOdt
    {
        private List<ModelTR> model;
        private int? width; // if not set, keep it is auto
        public ViewBodyOdt(List<ModelTR> model, int? width = null)
        {
            this.model = model;
            this.width = width;
        }
        public string render()
        {
            string table_width = width == null ? "" : string.Format("width={0}", width),
                   trs = "";
            if(width != null)
            {
                table_width = "";
                for (int i = 0; i < width; i++)
                {
                   
                    table_width += "<table:table-column table:style-name='TableColumn'/>";
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
                            attributes += string.Format("table:style-name='TableCellData' table:number-columns-spanned=\"{0}\" ", td.colspan);
                            for (int i = 1; i < td.colspan; i++)
                            {
                                table_cell += "<table:covered-table-cell/>";
                            }
                        }
                        if (td.className == "column")
                        {
                            attributes += string.Format("table:style-name='TitleCell'");
                            className = "TitleCellWord";
                        }
                        if (td.className == null)
                        {
                            attributes += string.Format("table:style-name='CellWord'");
                            className = "Word";
                        }
                        if (td.style == "TotalCell")
                        {
                            table_cell = "";
                            attributes = "";
                            var lastCellStyle = "table:style-name='CellWord'";
                            var lastTextStyle = "text:style-name='Word'";
                            var lastData = data;
                            attributes += string.Format("table:style-name='TotalCell' table:number-columns-spanned=\"{0}\" ", td.colspan);
                            for (int i = 1; i < td.colspan; i++)
                            {

                                table_cell += "<table:covered-table-cell/>";
                            }
                            table_cell += string.Format(lastRow, lastCellStyle, lastTextStyle, lastData);
                            data = "總比數";
                        }
                        if (td.className == "Title")
                        {
                            data = string.Format(titleRow, data);
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


        string template = @"<table:table table:style-name='Table'>
        <table:table-columns>{1}</table:table-columns>{0}</table:table>";

        string template_td = "<table:table-cell  {0}><text:p {1}>{2}</text:p></table:table-cell> {3}";
        string template_tr = "<table:table-row table:style-name='TableRow'>{1}</table:table-row>";
        string lastRow = "<table:table-cell  {0}><text:p {1}>{2}</text:p></table:table-cell>";
        string titleRow = "<text:span text:style-name='TitleWord'>{0}</text:span>";
        //0:style-name='TableCellData' table:number-columns-spanned=\"{0}\"  1:text:style-name="Title" 2:text:style-name="TitleWord" 3:data
    }
}

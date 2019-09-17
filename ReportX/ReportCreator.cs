using ReportX.Rep.Attributes;
using ReportX.Rep.Common;
using ReportX.Rep.Model;
using ReportX.Rep.View;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ReportX
{

    public class ReportCreator<T> where T : IReportX   
    {
        static T report { get; set; }


        public MemberInfo[] modeli;
        protected Type type { get; set; }
        protected string[] oldcols { get; set; }
        protected string[] newcols { get; set; }
        protected List<ModelTR> trs { get; }
        public string[] cols { get; set; }
        public string sheetName { get; set; }
        private ModelExcel excel;
        private ModelWord word;

        public ReportCreator(Type type)
        {
            excel = new ModelExcel();
            excel.style = new ViewStyle();

            word = new ModelWord();
            word.style = new ViewStyle();

            List<MemberInfo> list_cols = new List<MemberInfo>();
            modeli = type.GetMembers();
            foreach (var member in type.GetMembers())
            {
                Present attr = member.GetCustomAttribute<Present>();
                if (attr == null) continue;

                int MetadataToken = member.MetadataToken,
                    inserted_index = 0;

                // sory by MetadataToken (declaration)
                for (int i = 0; i < list_cols.Count; i++)
                {
                    inserted_index = i;
                    if (MetadataToken < list_cols[i].MetadataToken) break;
                    inserted_index = i + 1;
                }
                list_cols.Insert(inserted_index, member);
            }
            string[] str_cols = new string[list_cols.Count]; //取得標題數量

            for (int i = 0; i < list_cols.Count; i++)
                str_cols[i] = list_cols[i].GetCustomAttribute<Present>().getName();//取得標題名稱

            oldcols = str_cols; //舊的陣列
            cols = str_cols;
            
            excel.colNum = cols.Length;
            word.colNum = cols.Length;
            report =(T)Activator.CreateInstance(typeof(T), type);
        }
        public string render()
        {
            return report.render();
        }


        public void setDate(DateTime from, DateTime? to = null)
        {
            if (from == null) return;
            if (to == null) to = DateTime.Now;

            string date_start = Convert.ToDateTime(from).ToString("yyyy/MM/dd"),
                   date_end = Convert.ToDateTime(to).ToString("yyyy/MM/dd");
            report.appendFullRow(string.Format("{0} - {1}", date_start, date_end), null, "r-header-date");
        }
        public void setCreator(string creator)
        {
            report.setData(author: creator);
            report.appendFullRow(string.Format("製表人：{0}", creator), null, "r-header-secondary");
        }
        public void setTile(string title)
        {
            report.setData(sheetName : title);
            report.appendFullRow(title, null, "r-header-title");
        }
        public void setCreatedDate()
        {
            string now = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd hh:mm:tt");
            report.appendFullRow(string.Format("製表時間：{0}", now), null, "r-header-secondary");
        }

        public void setColumn()
        {
            ModelTR col = report.appendRow(cols);
            foreach (ModelTD td in col.tds)
                td.className = "column";
        }


        public void setData<T>(T[] data)
        {
            report.appendTable(data);
        }

        public void setsum<T>(T[] data) //總筆數
        {
            string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS
            report.appendRow(new { value = "總筆數", colspan = report.getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數

        }

        // 傳入欲顯示欄位標題 之陣列
        public void setcut(string[] cut)
        {
            report.changecut(cut);
        }
    }

}
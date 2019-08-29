using ReportX.Rep.Excel;
using ReportX.Rep.Integration;
using ReportX.Rep.Odf;
using ReportX.Rep.S5report;
using ReportX.Rep.Word;
using System;
using System.Data;

namespace ReportX
{
    public class Report
    {
        public ExcelReport excelResponse<T>(T[] data, string[] cols ,string title, DateTime starting, DateTime ending ,string Creator, bool end = false)
        {
            ExcelReport erp = new ExcelReport(typeof(T));
            
            if (cols.Length > 0)
            {
                erp.setcut(cols);
            }
            erp.setTile(title);
            erp.setDate(starting, ending);
            erp.setCreator(Creator);
            erp.setCreatedDate();
            erp.setColumn();
            erp.setData(data);

            if(end) //如果要顯示結算筆數 end =true;
            {
                erp.setsum(data);
            }
            return erp;
        }

        public WordReport WordResponse<T>(T[] data,string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false)
        {
            WordReport wrp = new WordReport(typeof(T));
            if (cols.Length > 0)
            {
                wrp.setcut(cols);
            }

            wrp.setTile(title);
            wrp.setDate(starting, ending);
            wrp.setCreator(Creator);
            wrp.setCreatedDate();
            wrp.setColumn();
            wrp.setData(data);

            if (end) //如果要顯示結算筆數 end =true;
            {
                wrp.setsum(data);
            }
            return wrp;   

        }
        public OdtReport OdtResponse<T>(T[] data, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false)
        {
            OdtReport orp = new OdtReport(typeof(T));
            if (cols.Length > 0)
            {
                orp.setcut(cols);
            }

            orp.setTile(title);
            orp.setDate(starting, ending);
            orp.setCreator(Creator);
            orp.setCreatedDate();
            orp.setColumn();
            orp.setData(data);

            if (end) //如果要顯示結算筆數 end =true;
            {
                orp.setsum(data);
            }
            return orp;

        }
        //綜合板
        public FileReport FileReport<T>(T[] data, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false)
        {
            FileReport file = new FileReport(typeof(T));
            if (cols.Length > 0)
            {
                file.setcut(cols);
            }

            file.setTile(title);
            file.setDate(starting, ending);
            file.setCreator(Creator);
            file.setCreatedDate();
            file.setColumn();
            file.setData(data);

            if (end) //如果要顯示結算筆數 end =true;
            {
                file.setsum(data);
            }
            return file;
        }
        public AmountReport AmountReport<T>(T[] data, string[] cols, string title,string dateTime,int sum_correct, int sum_wrong,  string Creator, bool end = false)
        {
            AmountReport file = new AmountReport(typeof(T));
            if (cols.Length > 0)
            {
                file.setcut(cols);
            }

            file.setTile(title);
            file.setCreator(Creator);
            file.setCreatedDate(dateTime);
            file.setColumn();
            file.setData(data);

            if (end) //如果要顯示結算筆數 end =true;
            {
                file.setsum(sum_correct, sum_wrong);
            }
            return file;
        }
    }
}

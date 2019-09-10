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
        public OdtReport OdtResponse<T>( T[] data, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false)
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
        public OdtReport OdtResponse(DataTable dtTable, string[] cols, string title, DateTime starting, DateTime ending, string Creator, bool end = false)
        {
            OdtReport orp = new OdtReport(dtTable);
            if (cols.Length > 0)
            {
                orp.setcut(cols);
            }

            orp.setTile(title);
            orp.setDate(starting, ending);
            orp.setCreator(Creator);
            orp.setCreatedDate();
            orp.setColumn();
            orp.setData(dtTable);

            if (end) //如果要顯示結算筆數 end =true;
            {
                orp.setsum(dtTable);
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
        //s5amount
        public AmountReport AmountReport<T>(T[] data, string[] cols, string title,string dateTime,  string Creator, bool end = false)
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
                file.setsum(data);
            }
            return file;
        }
        public KBStaticReport KBStaticReport<T>(T[] data, string[] cols, string title, string dateTime, string firstday, string lastdday, string Creator, bool end = false)
        {
            KBStaticReport file = new KBStaticReport(typeof(T));
            if (cols.Length > 0)
            {
                file.setcut(cols);
            }

            file.setTile(title);
            file.setCreator(Creator);
            file.setCreatedDate(dateTime);
            file.setCreatedDayRange(firstday, lastdday);
            file.setColumn();
            file.setData(data);

            if (end) //如果要顯示結算筆數 end =true;
            {
                file.setsum(data);
            }
            return file;
        }
        public SignStatusReport SignStatusReport<T>(T[] data, string[] cols, string title, string dateTime, string firstday, string lastdday, string Creator, bool end = false)
        {
            SignStatusReport file = new SignStatusReport(typeof(T));
            if (cols.Length > 0)
            {
                file.setcut(cols);
            }

            file.setTile(title);
            file.setCreator(Creator);
            file.setCreatedDate(dateTime);
            file.setCreatedDayRange(firstday, lastdday);
            file.setColumn();
            file.setSecondColumn();
            file.setData(data); ;

            if (end) //如果要顯示結算筆數 end =true;
            {
                file.setsum(data);
            }
            return file;
        }

    }
}

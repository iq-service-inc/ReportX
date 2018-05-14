using ReportX.Rep.Excel;
using ReportX.Rep.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ReportX
{
    public class Report
    {
        
        public ExcelReport excelResponse<T>(T[] data, string title, DateTime starting, DateTime ending ,string Creator)
        {
            ExcelReport erp = new ExcelReport(typeof(T));
            erp.setTile(title);
            erp.setDate(starting, ending);
            erp.setCreator(Creator);
            erp.setCreatedDate();
            erp.setColumn();
            erp.setData(data);

            return erp;

        }

        public WordReport WordResponse<T>(T[] data, string title, DateTime starting, DateTime ending, string Creator)
        {
            WordReport wrp = new WordReport(typeof(T));
            wrp.setTile(title);
            wrp.setDate(starting, ending);
            wrp.setCreator(Creator);
            wrp.setCreatedDate();
            wrp.setColumn();
            wrp.setData(data);
            
            return wrp;   

        }






    }
}

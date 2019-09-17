using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.Common
{
    public interface IReportX
    {
        string render(int? width = null);
        void changecut(string[] cut);
        void setCustomStyle(string css);
        ModelTR appendFullRow(string data, string trStyle = null, string className = null);
        ModelTR appendRow(params object[] data);
        void appendTable<T>(T[] data, string trStyle = null, string className = null);
        void appendTable(DataTable data, string trStyle = null, string className = null);
        void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null);
        int getColCount();
    }
}
    
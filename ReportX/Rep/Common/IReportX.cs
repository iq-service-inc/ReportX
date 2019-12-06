using ReportX.Rep.Model;
using System.Data;

namespace ReportX.Rep.Common
{
    public interface IReportX
    {
        string[] oldcols { get; set; }
        string[] newcols { get; set; }
        string[] cols { get; set; }
        string render(int? width = null);
        void changecut(string[] cut);
        void setCustomStyle(string css);
        ModelTR appendFullRow(string data, string trStyle = null, string className = null);
        ModelTR appendRow(params object[] data);
        void appendTable<T>(T[] data, string trStyle = null, string className = null);
        void appendTable(DataTable data, string trStyle = null, string className = null);
        void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null);
        int getColCount();
        void setCol<T>(T[] data);
        void setCol(DataTable data);
    }
}

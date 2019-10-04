using ReportXTests2.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportXTests2
{
    public class SampleData
    {
        public string[] ModelCol()
        {
            string[] cols = new string[4];
            cols[0] = "姓名";
            cols[1] = "資料";
            cols[2] = "ID";
            cols[3] = "電話";
            return cols;
        }
        public ModelEmployeeTicket[] ModelData()
        {
            ModelEmployeeTicket[] data = new ModelEmployeeTicket[5];
            for (int i = 5 - 1; i >= 0; i--)
            {
                ModelEmployeeTicket tmp = new ModelEmployeeTicket
                {
                    postpid = i + 1,
                    posttitle = "測試_" + i,
                    name = "SOL_" + i,
                    number = "123",
                    data = "data",
                    tel = "0923456789"
                };
                data[i] = tmp;
            }
            return data;
        }
        public DataTable Dtdata()
        {
            DataTable dtTable = new DataTable("dTable");
            DataColumn[] cols ={
            new DataColumn("ID",typeof(int)),
            new DataColumn("標題",typeof(string)),
            new DataColumn("姓名",typeof(string)),
            new DataColumn("電話",typeof(string)),
            new DataColumn("編號",typeof(string)),
            new DataColumn("資料",typeof(string))
            };
            dtTable.Columns.AddRange(cols);

            // 新增資料到DataTable
            for (int i = 0; i <= 4; i++)
            {
                var row = dtTable.NewRow();
                row["ID"] = i+1;
                row["標題"] = "測試 " + i.ToString();
                row["姓名"] = "SOL_" + i;
                row["編號"] = "123";
                row["資料"] = "data";
                row["電話"] = "0923456789";
                dtTable.Rows.Add(row);
            }
            return dtTable;
        }
    }
}

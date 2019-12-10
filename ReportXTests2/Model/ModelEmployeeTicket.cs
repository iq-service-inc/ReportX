using ReportX.Rep.Attributes;
using System;

namespace ReportXTests2.Model
{
    public class ModelEmployeeTicket
    {
        [Present("ID")]
        public Int64 postpid { get; set; }
        [Present("標題")]
        public string posttitle { get; set; }
        [Present("姓名")]
        public string name { get; set; }
        [Present("編號")]
        public string number{ get; set; }
        [Present("資料")]
        public string data { get; set; }
        [Present("電話")]
        public string tel { get; set; }
    }
}

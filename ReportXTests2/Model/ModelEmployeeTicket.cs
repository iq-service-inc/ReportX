using ReportX.Rep.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}

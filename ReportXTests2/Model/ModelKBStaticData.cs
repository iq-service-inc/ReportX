using ReportX.Rep.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReportXTests2.Model
{
    public class ModelKBStaticData
    {
        [Present("編號")]
        public int number { get; set; }
        [Present("知識目錄")]
        public string knowledge { get; set; }
        [Present("知識標題")]
        public string knowledgeTitle { get; set; }
        [Present("建立時間")]
        public string CreatedDate { get; set; }
        [Present("建立人員")]
        public string Creater { get; set; }
    }
}
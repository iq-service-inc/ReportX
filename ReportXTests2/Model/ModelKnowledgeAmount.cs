using ReportX.Rep.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReportXTests2.Model
{
   public class ModelKnowledgeAmount
    {

        [Present("順序")]
        public int sequence{ get; set; }
        [Present("知識目錄")]
        public string knowledge { get; set; }
        [Present("有效知識")]
        public int correctAmount { get; set; }
        [Present("無效知識")]
        public int wrongAmount { get; set; }
    }
}
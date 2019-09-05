using ReportX.Rep.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportXTests2.Model
{
    public class ModelSignStatusData
    {
        [Present("順序")]
        public int sequence { get; set; }
        [Present("審核起始日")]
        public string reviewDate { get; set; }
        [Present("新增知識")]
        public string addknowledge { get; set; }
        [Present("修改知識")]
        public string updateknowledge { get; set; }
        [Present("刪除知識")]
        public string deleteknowledge { get; set; }

        [Present("新增知識審核")]
        public int addknowledgeReview { get; set; }
        [Present("新增知識生效")]
        public int addknowledgeWork { get; set; }
        [Present("新增知識退件")]
        public int addknowledgeBack { get; set; }
        [Present("修改知識審核")]
        public int updateknowledgeReview { get; set; }
        [Present("修改知識生效")]
        public int updateknowledgeWork { get; set; }
        [Present("修改知識退件")]
        public int updateknowledgeBack { get; set; }
        [Present("刪除知識審核")]
        public int deleteknowledgeReview { get; set; }
        [Present("刪除知識生效")]
        public int deleteknowledgeWork { get; set; }
        [Present("刪除知識退件")]
        public int deleteknowledgeBack { get; set; }
        [Present("合計")]
        public int total { get; set; }
    }
}

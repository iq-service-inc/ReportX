using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyReportX.Rep.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class Present : Attribute
    {
        private string name;
        public Present(string name)
        {
            this.name = name;
        }
        public string getName()
        {
            return name;
        }
    }
}

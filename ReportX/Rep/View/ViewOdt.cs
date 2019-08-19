using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
   public class ViewOdt
    {
        private ModelOdt m;

        public ViewOdt(ModelOdt model)
        {
            m = model;
        }

        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(m.author, m.company, m.sheetName, style, body);

        }

        

    }
}

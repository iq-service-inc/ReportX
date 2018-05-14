using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
{
    public class ViewExcel
    {
        private ModelExcel m;

        public ViewExcel(ModelExcel model)
        {
            m = model;
        }

        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(template, m.author, m.company, m.sheetName, style, body);

        }

        string template = @"
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>
            <head>
                <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
                <meta name=ProgId content=Excel.Sheet>
                <meta name=Generator content='Microsoft Excel 11'>

                <!--[if gte mso 9]><xml>
                <o:DocumentProperties>
                <o:Author>{0}</o:Author>
                <o:Company>{1}</o:Company>
                </o:DocumentProperties>
                </xml><![endif]-->

                <!--[if gte mso 9]><xml>
                <x:ExcelWorkbook>
                <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                <x:Name>{2}</x:Name>
                <x:WorksheetOptions>
                <x:DefaultRowHeight>200</x:DefaultRowHeight>
                <x:Selected/>
                <x:ProtectContents>False</x:ProtectContents>
                <x:ProtectObjects>False</x:ProtectObjects>
                <x:ProtectScenarios>False</x:ProtectScenarios>
                </x:WorksheetOptions>
                </x:ExcelWorksheet>
                </x:ExcelWorksheets>
                <x:WindowHeight>8160</x:WindowHeight>
                <x:WindowWidth>11715</x:WindowWidth>
                <x:WindowTopX>240</x:WindowTopX>
                <x:WindowTopY>75</x:WindowTopY>
                <x:ProtectStructure>False</x:ProtectStructure>
                <x:ProtectWindows>False</x:ProtectWindows>
                </x:ExcelWorkbook>
                </xml><![endif]-->          
                {3}
            </head>
            <body>
                {4} 
            </body>
            </html>";
    }

}

using MyReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyReportX.Rep.View
{
    public class ViewWord
    {
        private ModelWord m;

        public ViewWord(ModelWord model)
        {
            m = model;
        }

        public string render()
        {
            string style = m.style.render(),
                   body = m.body.render();

            // more coustom code here
            // ...

            return string.Format(word, m.author, m.company, m.sheetName, style, body);

        }
        string word = @"
                <html xmlns:v='urn:schemas-microsoft-com:vml'
                xmlns:o='urn:schemas-microsoft-com:office:office'
                xmlns:w='urn:schemas-microsoft-com:office:word'
                xmlns:m='http://schemas.microsoft.com/office/2004/12/omml'
                xmlns='http://www.w3.org/TR/REC-html40'>

                <head>
                <meta http-equiv=Content-Type content = 'text/html; charset=utf-8'>
                <meta name = ProgId content = Word.Document>
                <meta name = Generator content = 'Microsoft Word 15'>
                <meta name = Originator content = 'Microsoft Word 15'>
                <link rel = File-List href = 'Doc1.files/filelist.xml'>

                 <!--[if gte mso 9]><xml>
                 <o:DocumentProperties>
                  <o:Author>{0}</o:Author>
                  <o:Template>Normal</o:Template>
                  <o:LastAuthor>{1}</o:LastAuthor>
                  <o:Revision>2</o:Revision>
                  <o:TotalTime>{2}</o:TotalTime>
                  <o:Created>2018-05-08T05:37:00Z</o:Created>
                  <o:LastSaved>2018-05-08T05:37:00Z</o:LastSaved>
                  <o:Pages>1</o:Pages>
                  <o:Version>16.00</o:Version>
                 </o:DocumentProperties>
                 <o:OfficeDocumentSettings>
                  <o:AllowPNG/>
                 </o:OfficeDocumentSettings>
                </xml><![endif]-->
                    {3}
                <!--[if gte mso 9]><xml>
                 <w:WordDocument>
                  <w:TrackMoves>false</w:TrackMoves>
                  <w:TrackFormatting/>
                  <w:PunctuationKerning/>
                  <w:DisplayHorizontalDrawingGridEvery>0</w:DisplayHorizontalDrawingGridEvery>
                  <w:DisplayVerticalDrawingGridEvery>2</w:DisplayVerticalDrawingGridEvery>
                  <w:ValidateAgainstSchemas/>
                  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
                  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
                  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
                  <w:DoNotPromoteQF/>
                  <w:LidThemeOther>EN-US</w:LidThemeOther>
                  <w:LidThemeAsian>ZH-TW</w:LidThemeAsian>
                  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
                  <w:Compatibility>
                   <w:SpaceForUL/>
                   <w:BalanceSingleByteDoubleByteWidth/>
                   <w:DoNotLeaveBackslashAlone/>
                   <w:ULTrailSpace/>
                   <w:DoNotExpandShiftReturn/>
                   <w:AdjustLineHeightInTable/>
                   <w:BreakWrappedTables/>
                   <w:SnapToGridInCell/>
                   <w:WrapTextWithPunct/>
                   <w:UseAsianBreakRules/>
                   <w:DontGrowAutofit/>
                   <w:SplitPgBreakAndParaMark/>
                   <w:EnableOpenTypeKerning/>
                   <w:DontFlipMirrorIndents/>
                   <w:OverrideTableStyleHps/>
                   <w:UseFELayout/>
                  </w:Compatibility>
                  <m:mathPr>
                   <m:mathFont m:val='Cambria Math'/>
                   <m:brkBin m:val='before'/>
                   <m:brkBinSub m:val='&#45;-'/>
                   <m:smallFrac m:val='off'/>
                   <m:dispDef/>
                   <m:lMargin m:val='0'/>
                   <m:rMargin m:val='0'/>
                   <m:defJc m:val='centerGroup'/>
                   <m:wrapIndent m:val='1440'/>
                   <m:intLim m:val='subSup'/>
                   <m:naryLim m:val='undOvr'/>
                  </m:mathPr></w:WordDocument>
                </xml><![endif]-->
                    {3}
                </head>
                    <body>
                    {4}
                   </body>
                </html>";

    }
}

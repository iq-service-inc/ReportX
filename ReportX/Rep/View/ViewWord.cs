using ReportX.Rep.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportX.Rep.View
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

            return string.Format(wordtest, m.author, m.company, m.sheetName, style, body);

        }

        string wordtest = @"
            <html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns:m='http://schemas.microsoft.com/office/2004/12/omml' xmlns='http://www.w3.org/TR/REC-html40'>
            <head>
            <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
            <meta name = 'ProgId' content='Word.Document'>
            <meta name = 'Generator' content='Microsoft Word 15'>
            <meta name = 'Originator' content='Microsoft Word 15'>
            <link rel = 'File-List' href='1.files/filelist.xml'>

            <!--[if gte mso 9]><xml>
             <o:DocumentProperties>
              <o:Author>{0}</o:Author>
              <o:Template>Normal</o:Template>
              <o:LastAuthor>{0}</o:LastAuthor>
              <o:Revision>1</o:Revision>
              <o:TotalTime>8</o:TotalTime>
              <o:Created>{1}</o:Created>
              <o:LastSaved>{2}</o:LastSaved>
              <o:Pages>1</o:Pages>
              <o:Words>1</o:Words>
              <o:Characters>10</o:Characters>
              <o:Lines>1</o:Lines>
              <o:Paragraphs>1</o:Paragraphs>
              <o:CharactersWithSpaces>10</o:CharactersWithSpaces>
              <o:Version>16.00</o:Version>
             </o:DocumentProperties>
             <o:OfficeDocumentSettings>
              <o:AllowPNG/>
             </o:OfficeDocumentSettings>
            </xml><![endif]-->
            <!--[if gte mso 9]><xml>
             <w:WordDocument>
              <w:SpellingState>Clean</w:SpellingState>
              <w:GrammarState>Clean</w:GrammarState>
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
            <body lang='ZH-TW' style='tab-interval:24.0pt;text-justify-trim:punctuation'>
            {4}
            </body>
            </html>";

    }
}

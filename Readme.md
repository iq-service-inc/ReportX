# ReportX
ReportX 可以用簡單的方法，快速建立 Word 與 Excel 報表。更棒的是，它的輸出結果也可以在網頁中直接呈現

## Installation

**Package Manager**

```
PM> Install-Package ReportX -Version 1.2.0
```

## System requirement

* `v1.2.0` 開始的版本僅支援 .NET Framework 4.5 以上

## API Reference

* Report： 
* ExcelReport： 
* WordReport： 


## Customized

* `v1.2.0`
  * 可以自訂報表規則，產生Excel報表範例如下：  

```csharp

ExcelReport Exx = new ExcelReport(typeof(data));  //data 為資料陣列

Exx.setTile('設置標題');  
Exx.setDate(Convert.ToDateTime('開始時間'), Convert.ToDateTime('結束時間')); 
Exx.setCreatedDate();  //製表時間
Exx.setColumn(); //建立表格屬性
Exx.setData(data); //匯入資料
            
//統計資料數
Exx.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length);
            
//產生報表
string output = Exx.render();
            
```

## Default

* `v1.2.0`
  * 可以使用內建報表規則，使用範例如下：  

```csharp

Report s = new Report(); //使用Report 方法

//帶入參數('資料','標題','開始時間','結束時間','製表人')
ExcelReport myca = s.excelResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");
//統計資料數
myca.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);
//產生報表
string output = Exx.render();
//輸出報表

            
```

## Output 

* 用 File.AppendAllText 方法 把檔案保存到電腦

```csharp

File.AppendAllText(output, ".doc"); //另存為Word檔
File.AppendAllText(output, ".xls"); //另存為Excel檔


```
## use ReportX with ZapLib
* 使用 Zaplib 中的 ExtApiHelper 套件，呼叫 getAttachmentResponse()
* 下載預設為 .xls

```csharp

ExtApiHelper api = new ExtApiHelper(this); 

return api.getAttachmentResponse(資料,檔名);

```

即可下載excel檔案

## License

   Copyright 2018 LinSol

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
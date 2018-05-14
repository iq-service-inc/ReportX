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


## Default

* `v1.2.0` 可以使用內建報表規則，使用範例如下：  

```csharp

Report s = new Report(); 

//帶入參數產生Excel報表 (資料 , 標題 , 開始時間 , 結束時間 , 製表人 )
ExcelReport myex = s.excelResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");
//帶入參數產生Word報表 (資料 , 標題 , 開始時間 , 結束時間 , 製表人 )
WordReport mywd = s.wordResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");

//統計資料數
myex.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);
mywd.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);

//產生報表
string Outexcel = myex.render();
string Outword = mywd.render();

//另存檔案
File.AppendAllText(Outexcel, ".xls"); //另存為Excel檔
File.AppendAllText(Outword, ".doc");  //另存為Word檔

            
```

## Customized Excel

* `v1.2.0` 自訂規則製作成Excel檔，使用範例如下：

```csharp

ExcelReport Exx = new ExcelReport(typeof("data"));  //data 為資料陣列

Exx.setTile("設置標題");  
Exx.setDate(Convert.ToDateTime("開始時間"), Convert.ToDateTime("結束時間")); 
Exx.setCreatedDate();  //製表時間
Exx.setColumn();       //建立表格屬性
Exx.setData(data);     //匯入資料
            
//統計資料數
Exx.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length);
            
//產生報表
string output = Exx.render();

//另存Excel 報表
File.AppendAllText(output, ".xls"); 

            
```

## Customized Word

* `v1.2.0` 自訂規則也可以製作成Word檔，使用範例如下：

```csharp

WordReport wrd = new WordReport(typeof("data"));  //data 為資料陣列

wrd.setTile("設置標題");  
wrd.setDate(Convert.ToDateTime("開始時間"), Convert.ToDateTime("結束時間")); 
wrd.setCreatedDate();  //製表時間
wrd.setColumn();       //建立表格屬性
wrd.setData(data);     //匯入資料
            
//統計資料數
wrd.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length);
            
//產生報表
string output = wrd.render();

//另存Word 報表
File.AppendAllText(output, ".doc"); 

            
```


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
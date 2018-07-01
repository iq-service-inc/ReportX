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

## Default Model
* 以下範例資料皆使用此Model：

```csharp
using ReportX.Rep.Attributes;

namespace TEST.Models
{
    public class ModelGO
    {
        [Present("編號")]
        public string postpid { get; set; }
        [Present("標題")]
        public string posttitle { get; set; }
        [Present("作者編號")]
        public string authormid { get; set; }
        [Present("員工編號")]
        public string employeeno { get; set; }
    }
}
```

## Default

* `v1.2.0` 使用內建規則報表，使用範例如下：  

```csharp
//範例模型
ModelGO[] data = new ModelGO[20];

for (int a = 0; a < 20; a++)
{
    data[a] = new ModelGO
    {
        postpid = "ID" + a,
        posttitle = "標題" + a,
        authormid = "編號" + a,
        employeeno = "員工" + a
    };
}

string lastRowStyle = "background-color:#DDD;-webkit-print-color-adjust: exact;"; //預設CSS

Report s = new Report(); 

//帶入參數產生Excel 報表 (資料 , 標題 , 開始時間 , 結束時間 , 製表人 )
ExcelReport myex = s.excelResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");
myex.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);
string Outexcel = myex.render();
File.AppendAllText("路徑+檔案.xls", Outexcel); //另存為Excel檔


//帶入參數產生Word 報表
WordReport mywd = s.wordResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");
mywd.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);
string Outword = mywd.render();
File.AppendAllText("路徑+檔案.doc", Outexcel);  //另存為Word檔
```

## Customized Excel

* `v1.2.0` 自訂規則製作成Excel檔，使用範例如下：

```csharp
//範例模型
ModelGO[] data = new ModelGO[20];

for (int a = 0; a < 20; a++)
{
    data[a] = new ModelGO
    {
        postpid = "ID" + a,
        posttitle = "標題" + a,
        authormid = "編號" + a,
        employeeno = "員工" + a
    };
}

ExcelReport Exx = new ExcelReport(typeof(data));  //data 為資料陣列

Exx.setTile("設置標題");  
Exx.setDate(Convert.ToDateTime("開始時間"), Convert.ToDateTime("結束時間")); 
Exx.setCreatedDate();  //製表時間
Exx.setColumn();       //建立表格屬性
Exx.setData(data);     //匯入資料
            
Exx.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length);//統計資料數
string output = Exx.render();//產生報表
File.AppendAllText(output, ".xls"); //另存Excel 報表
```

## Customized Word

* `v1.2.0` 自訂規則也可以製作成Word檔，使用範例如下：

```csharp
//範例模型
ModelGO[] data = new ModelGO[20];

for (int a = 0; a < 20; a++)
{
    data[a] = new ModelGO
    {
        postpid = "ID" + a,
        posttitle = "標題" + a,
        authormid = "編號" + a,
        employeeno = "員工" + a
    };
}

WordReport wrd = new WordReport(typeof(data));  //data 為資料陣列

wrd.setTile("設置標題");  
wrd.setDate(Convert.ToDateTime("開始時間"), Convert.ToDateTime("結束時間")); 
wrd.setCreatedDate();  //製表時間
wrd.setColumn();       //建立表格屬性
wrd.setData(data);     //匯入資料
            
wrd.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length); //統計資料數
string output = wrd.render();//產生報表
File.AppendAllText(output, ".doc"); //另存Word 報表
```
## Preview

![image](http://192.168.1.136/SideProject/ReportX/raw/master/EX.PNG)
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
# ReportX 3
ReportX 3 可以用簡單的方法，快速建立 Word, Excel, Odt 與 Ods 報表。  
你可以使用內建的快速產生器建立預設報表，也可以照自己的意思刻出客製化報表。

## Installation

**Package Manager**

```text
PM> Install-Package ReportX -Version 3.0.0
```

## System requirement

* `v1.2.0` 開始的版本僅支援 .NET Framework 4.5 以上

## Getting Started

快速建立一支 Word 報表，使用 `ReportCreator<T>` 產生器建立，將 `T` 帶入 `Word`, `Excel`, `Odt` 或 `Ods` 就可以產生相應格式的報表。

```csharp

// 資料
ModelEmployeeTicket[] data = new ModelEmployeeTicket[] {
    new ModelEmployeeTicket(){ postpid=10, name="zap"},
    new ModelEmployeeTicket(){ postpid=11, name="jack"},
    new ModelEmployeeTicket(){ postpid=12, name="peter"},
};

// 要顯示的欄位 (不一定要全部欄位都顯示)
string[] cols = new string[] { "ID", "姓名" };

// 報表標題
string title = "測試報表";

// 報表資料的時間範圍
DateTime date_from = DateTime.Now.AddDays(-1);
DateTime date_to = DateTime.Now;

// 建立報表人
string creator = "Administrator";

// 是否顯示資料總筆數
bool showTotal = true;

// 建立 Word 報表
ReportCreator<Word> report = new ReportCreator<Word>();
report.setInfo(data, cols, title, date_from, date_to, creator, showTotal);

// 報表結果字串，可以直接存檔成 .doc 即可瀏覽
string word = report.render();
```

其中，關於 `data` 資料模型 `ModelEmployeeTicket` 可以自行定義，可以加上 `[Present("顯示名稱")]` 標籤來設定該欄位要顯示的欄位名稱，請參考以下範例：

```csharp
public class ModelEmployeeTicket
{
    [Present("ID")]
    public Int64 postpid { get; set; }
    [Present("姓名")]
    public string name { get; set; }
    [Present("電話")]
    public string tel { get; set; }
}
```

使用內建的報表存檔類別，將報表存檔

```csharp
// 將報表物件傳入
ReportFile rf = new ReportFile(report.report);
string fileName = "我的報表";
// 儲存報表，將回傳報表的儲存路徑 (報表將會被存放在暫存區，你可以自行再搬移)
string path = rf.saveFile(fileName);
```

## More Examples

以下展示更多範例，您可自行參閱最符合需求的案例參考

> 部分範例將使用 [Getting Started](#Getting-Started) 章節中的 `ModelEmployeeTicket` 資料模型


### 產生 Excel 報表

```csharp
```



## License

   Copyright 2020 Zap Lin

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
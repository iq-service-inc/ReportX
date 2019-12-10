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




## `ReportCreator` 建立標準規格報表

ReportX 提供了一個標準報表產生器  `ReportCreator<T>`，它包含了：
* **標題**
* **資料的時間範圍**
* **建立報表人**
* **報表建立時間**
* **資料表格 (可設定欲顯示的資料欄位)**


> **❗ 注意：**
> * 只有 `Word` 與 `Excel` 的產生的報表結果支援網頁顯示 (屬於 HTML)，`Odt` 與 `Ods` 不支援 (特殊格式的 XML)
> * `Odt` 與 `Ods` 因為需要將 meta 檔案與報表內容檔進行 zip 壓縮後才可正常瀏覽，因此需要先存成實體檔案才可使用


將 `<T>` 帶入 `Word`, `Excel`, `Odt` 或 `Ods` 就可以產生相應格式的報表：


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

使用內建的實體報表存檔工具 `ReportFile` ，將報表存成實際檔案

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





### 產生 OpenOffice 報表

OpenOffice 報表是由以下結構所構成：

* OpenOffice 檔 (zip 壓縮)
    * `META-INF`：設定檔存放資料夾
        * `manifest.xml`：OpenOffice 文件設定檔
    * `content.xml`：報表內容

因此，如果要瀏覽 OpenOffice 檔案，需要先將報表使用以上結構組成後 zip 壓縮，但 ReportX 提供了更簡易的方法：

```csharp
ReportCreator<Ods> report = new ReportCreator<Ods>();
report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
// 將報表儲存
ReportFile rf = new ReportFile(report.report);
string fileName = "My_OpenOffice_Report";
string path = rf.saveFile(fileName); // 回傳存放路徑 (可再自行移動)
```




### 將多個 Excel 合併成一個

目前支援 Microsoft Excel 將多個報表整合成一個，使用 `MultiExcelBundler` 類別

```csharp
// 建立第一張 Excel
ReportCreator<Excel> report1 = new ReportCreator<Excel>();
report1.setInfo(data, cols, "第一個Excel", DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

// 建立第二張 Excel 
ReportCreator<Excel> report2 = new ReportCreator<Excel>();
report2.setInfo(data, cols, "第二個Excel", DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

// 綁定兩個 Excel 
MultiExcelBundler bundler = new MultiExcelBundler();
bundler.addExcel(report1.report);
bundler.addExcel(report2.report);

// 儲存成實體檔案 (將 bundler 帶入)
ReportFile rf = new ReportFile(bundler);
string path = rf.saveFile(fileName);
```




### 支援 `DataTable` 資料輸入

在 [Getting Started](#Getting-Started) 章節中使用 `ModelEmployeeTicket` 資料模型做為資料儲存的容器。  
此外，也可使用 `DataTable` 作為報表的資料輸入，使用上與資料模型並無差異，範例如下：

```csharp
DataTable data = new DataTable("dTable");
// 設定欄位
DataColumn[] table_column ={
    new DataColumn("ID",typeof(int)),
    new DataColumn("姓名",typeof(string)),
    new DataColumn("電話",typeof(string))
};
data.Columns.AddRange(table_column);

// 填充資料
for (int i = 0; i <= 4; i++)
{
    var row = data.NewRow();
    row["ID"] = i+1;
    row["姓名"] = "SOL";
    row["電話"] = "0923456789";
    data.Rows.Add(row);
}

// 使用 DataTable 做為資料輸入
ReportCreator<Word> report = new ReportCreator<Word>();
report.setInfo(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
string word = report.render();
```

### 完全客製化報表

如果 `ReportCreator` 預設標準的報表格式不符合使用需求，如果要自訂報表，可以參考以下範例：

```csharp
// 舉例：客製化 Word 報表的內容
Word report = new Word();

// 設定欄位資訊，將 data 的欄位填入
report.setCol(data);

// 設定要顯示在報表上的欄位
string[] cols = new string[] { "ID", "姓名" };
report.changecut(cols);

// 增加一個滿版橫列 (橫跨所有欄位)
string className = "r-header-title";
report.appendFullRow("增加一個 Title", null, className);

// 填充客製化樣式設定 (Word, Excel 使用 CSS；Odt, Ods 使用 XML)
string customOfficeCSS = @"
.r-header-title{
    font-size: 22px;
    font-weight: bold;
}"
report.setCustomStyle(customOfficeCSS);

// 使用 ModelTR 與 ModelTD 進行組裝，類似以下效果
// <tr>
//      <td class='column'>資料1</td> 
//      <td class='column'>資料2</td>  
// </tr>
ModelTR col = report.appendRow(new string[] { "資料1", "資料2" });
// 針對每個 td 設定 className 或是 style
foreach (ModelTD td in col.tds) td.className = "column";

// 將 data 資料填充到報表中
report.appendTable(data);

// 畫出自定的報表
string res = report.render();
```

❗ 注意，`Odt` 與 `Ods` 的樣式設定不是 CSS，他是由 XML 結構組成的設定，範例如下：

```xml
<office:automatic-styles>
    <style:style style:name='ColumnWord' style:family='table-cell' style:parent-style-name='Default' style:data-style-name='N0'>
        <style:text-properties fo:color='#FFFFFF' style:font-name='微軟正黑體' style:font-name-asian='微軟正黑體' style:font-name-complex='微軟正黑體'/>
    </style:style>
    <style:page-layout style:name='pm1'>
        <style:page-layout-properties fo:margin-top='0.5in' fo:margin-bottom='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' style:print-orientation='portrait' style:print-page-order='ttb' style:first-page-number='continue' style:scale-to='100%' style:table-centering='none' style:print='objects charts drawings'/>
        <style:header-style>
            <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-bottom='0in'/>
        </style:header-style>
        <style:footer-style>
            <style:header-footer-properties fo:min-height='0.5in' fo:margin-left='0.75in' fo:margin-right='0.75in' fo:margin-top='0in'/>
        </style:footer-style>
    </style:page-layout>
</office:automatic-styles>

<office:master-styles>
    <style:master-page style:name='mp1' style:page-layout-name='pm1'>
        <style:header/>
        <style:header-left style:display='false'/>
        <style:footer/>
        <style:footer-left style:display='false'/>
    </style:master-page>
</office:master-styles>
```

## API Reference

ReportX API 參考

### `ReportCreator<T>` Class

快速建立標準報表物件

#### 成員

* `T report`：報表物件 (`T` 可為 `Word`,`Excel`,`Ods`,`Odt`)

#### 建構子

* `ReportCreator()`

#### 方法

* `void setInfo(R[] data, string[] cols, string title, DateTime from, DateTime? to = null, string creator = null, bool showTotal = false)`：設定報表資訊
    * `data`：資料
    * `cols`：欲顯示的欄位
    * `title`：標題
    * `from`：資料開始時間
    * `to`：資料結束時間
    * `creator`：報表建立人
    * `showTotal`：是否顯示資料總數

* `string render()`：畫出目前的報表，回傳報表字串結果


### `ReportFile` Class

將報表字串存檔成實體檔案專用類別，支援 IReport 與 MultiExcelCreator 兩種規格的檔案產生

#### 成員

無

#### 建構子

* `ReportFile(IReportX report)`
    * `report`：為 `Word`,`Excel`,`Ods`,`Odt` 其中一種 class
* `ReportFile(MultiExcelBundler excel_creator)`
    * `excel_creator`：多個 excel 合成好的 class 

#### 方法

* `string saveFile(string name, int? width = null)`：將報表儲存成實體檔案，並回傳儲存路徑
    * `fileName`：報表名稱(不用副檔名)
    * `width`：寬度
* `void deleteReportFile()`：如果報表已經不需要再使用，則可以呼叫此方法刪除檔案，否則需要自行刪除


### `MultiExcelBundler` Class

#### 成員

無

#### 建構子

* `MultiExcelBundler()`

#### 方法

* `void addExcel(Excel report)`：添加 Excel 報表
    * `report`：Excel 報表
* `string render(int? width = null)`：將多個 Excel 綁定成一個，並生成新的內容字串
    * `width`：寬度


### `Excel` Class

#### 成員

* 

#### 建構子

* 

#### 方法

* 




### `` Class

#### 成員

* 

#### 建構子

* 

#### 方法

* 




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
# 📜 ReportX V3 🎉🎉

全新 ReportX V3 可以用簡單的方法，快速建立 Word, Excel, Odt 與 Ods 報表。  
你可以使用內建的快速產生器建立預設報表，也可以照自己的意思刻出客製化報表。


## Installation

**Package Manager**

```text
PM> Install-Package ReportX -Version 3.0.0
```

**Nuget**

[ReportX 3.0.0](https://www.nuget.org/packages/ReportX/3.0.0)

## System requirement
* `v1.2.0` 開始的版本僅支援 .NET Framework 4.5 以上


## 🔰 基本使用範例

### 👍 `ReportCreator` 建立標準規格報表

ReportX 提供了一個標準報表產生器  `ReportCreator<T>`，它包含了：
* **標題**
* **資料的時間範圍**
* **建立報表人**
* **報表建立時間**
* **資料表格 (可設定欲顯示的資料欄位)**

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

加上 `[Present("顯示名稱")]` 屬性(Attribute) 設定欄位顯示名稱：

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
> 資料模型 `ModelEmployeeTicket` 可以自行定義


使用報表存檔類別 `ReportFile` ，將報表存成實際檔案

```csharp
// 將報表物件傳入
ReportFile rf = new ReportFile(report.report);
string fileName = "我的報表";
// 儲存報表，將回傳報表的儲存路徑 (報表將會被存放在暫存區，你可以自行再搬移)
string path = rf.saveFile(fileName);
```

**⚠️注意：**
* 只有 `Word` 與 `Excel` 的產生的報表結果支援網頁顯示 (屬於 HTML)，`Odt` 與 `Ods` 不支援 (特殊格式的 XML)
* `Odt` 與 `Ods` 因為需要將 meta 檔案與報表內容檔進行 zip 壓縮後才可正常瀏覽，因此需要先存成實體檔案才可使用
  

## 🔰 More Examples

以下展示更多範例，您可自行參閱最符合需求的案例參考

* [產生 OpenOffice 報表](#產生-openoffice-報表)
* [將多個 Excel 合併成一個](#將多個-excel-合併成一個)
* [支援 `DataTable` 資料輸入](#支援-datatable-資料輸入)
* [完全客製化報表](#完全客製化報表)

> 部分範例將使用 [Getting Started](#Getting-Started) 章節中的 `ModelEmployeeTicket` 資料模型

-------------------------------

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

-------------------------------


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

-------------------------------


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


-------------------------------

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

ReportX API 參考一覽

* [ReportCreator<T>](#reportcreatort-class)
* [ReportFile](#reportfile-class)
* [MultiExcelBundler](#multiexcelbundler-class)
* [Word](#word-class)
* [Excel](#excel-class)
* [Odt](#odt-class)
* [Ods](#ods-class)
* [AbsOffice 與 AbsOpenOffice Class](#absoffice-與-absopenoffice-class)
* [IReportX](#ireportx-interface)

----------------------------------------

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

-------------------------------

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

-------------------------------

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

-------------------------------
### `Word` Class

∟ 繼承：[`AbsOffice`](#absoffice-與-absopenoffice-class)  
Microsoft Office Word 底層操作類別
#### 成員

同 [IReportX](#absoffice-與-absopenoffice-class) class 介紹

#### 建構子

* `Word()`

#### 方法

同 [IReportX](#absoffice-與-absopenoffice-class) class 介紹

-------------------------------
### `Excel` Class

∟ 繼承：[`AbsOffice`](#absoffice-與-absopenoffice-class)  
Microsoft Office Excel 底層操作類別

#### 成員

同 [IReportX](#absoffice-與-absopenoffice-class) class 介紹

#### 建構子

* `Excel()`

#### 方法

同 [IReportX](#absoffice-與-absopenoffice-class) class 介紹

-------------------------------

### `Odt` Class

∟ 繼承：[`AbsOpenOffice`](#absoffice-與-absopenoffice-class)  
OpenOffice Odt 底層操作類別

#### 成員

* `string meta`：Ods file 專用 Meta 宣告，用於 META-INF 檔案建立時填入
其餘同 [AbsOpenOffice](#absoffice-與-absopenoffice-class) class 介紹

#### 建構子

* `Odt()`  

#### 方法

同 [AbsOpenOffice](#absoffice-與-absopenoffice-class) class 介紹

-------------------------------
### `Ods` Class

∟ 繼承：[`AbsOpenOffice`](#absoffice-與-absopenoffice-class)  
OpenOffice Ods 底層操作類別

#### 成員

* `string meta`：Ods file 專用 Meta 宣告，用於 META-INF 檔案建立時填入
其餘同 [AbsOpenOffice](#absoffice-與-absopenoffice-class) class 介紹

#### 建構子

* `Ods()`

#### 方法

同 [AbsOpenOffice](#absoffice-與-absopenoffice-class) class 介紹

-------------------------------

### `AbsOffice` 與 `AbsOpenOffice` Class
∟ 繼承：`IReport`   
定義 Office 與 OpenOffice 相關功能的抽象類別 

#### 成員

* `abstract string[] oldcols`：原欄位位資訊
* `abstract string[] newcols`：過濾後欄位資訊
* `abstract string[] cols`：資料欄位資訊
* `abstract List<ModelTR> trs`：報表每一列的詳細資訊

#### 建構子

抽象類別無法實例化

#### 方法

* `virtual string render(int? width = null)`：畫出報表，並回傳結果字串
    * `width`：寬度
* `abstract void changecut(string[] cut)`：過濾欲顯示欄位
    * `cut`：欲顯示的欄位陣列
* `abstract void setCustomStyle(string css)`：設定客製化 CSS 樣式
* `abstract ModelTR appendFullRow(string data, string trStyle = null, string className = null)`：新增一個滿版列(橫跨所有欄)
    * `data`：該列要顯示的內容
    * `trStyle`：該列的自訂樣式
    * `className`：該列的 className
* `ModelTR appendRow(params object[] data)`：新增一列，並在該列中填充數個欄位
    * `data`：每一個欄位的設定，必須是一個陣列。其中，`object` 的規格如下：
        * `object data`：要顯示的資料
        * `int colspan`：合併幾個欄，預設 1 無合併
        * `int rowspan`：合併幾個列，預設 1 無合併
        * `string fontSize`：字體大小
        * `string align`：對齊設定 (center, left, right)
        * `bool bold`：是否粗體
        * `string bgcolor`：背景顏色
        * `string style`：自訂樣式
        * `string className`：className
* `void appendTable<T>(T[] data, string trStyle = null, string className = null)`：新增一個 Table 資料
    * `data`：表格資料 (資料模型陣列)
    * `trStyle`：每列樣式
    * `className`：每列 className
* `void appendTable(DataTable data, string trStyle = null, string className = null)`：新增一個 Table 資料
    * `data`：表格資料 (DataTable)
    * `trStyle`：每列樣式
    * `className`：每列 className
* `void setCol<T>(T[] data)`：設定資料欄位資訊 (呼叫 `setReportColNum`)
* `void setCol(DataTable data)`：設定資料欄位資訊 (呼叫 `setReportColNum`)
* `abstract void setReportColNum()`：設定欄位數量
* `abstract void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)`：設定報表背景資訊
    * `author`：作者名稱
    * `company`：公司名稱
    * `sheetName`：報表名稱
    * `dateTime`：建立時間
    * `dateRange`：報表資料時間範圍

-------------------------------

### `IReportX` interface

基底報表規格介面

* `string[] oldcols`：定義原欄位位資訊
* `string[] newcols`：定義過濾後欄位資訊
* `string[] cols`：定義資料欄位資訊
* `string render(int? width = null)`：定義畫出報表，並回傳結果字串
* `void changecut(string[] cut)`：定義過濾欲顯示欄位
* `void setCustomStyle(string css)`：定義設定客製化 CSS 樣式
* `ModelTR appendFullRow(string data, string trStyle = null, string className = null)`：定義新增一個滿版列(橫跨所有欄)
* `ModelTR appendRow(params object[] data)`：定義新增一列，並在該列中填充數個欄位
* `void appendTable<T>(T[] data, string trStyle = null, string className = null)`：定義新增一個 Table 資料
* `void appendTable(DataTable data, string trStyle = null, string className = null)`：定義新增一個 Table 資料
* `void setData(string author = null, string company = null, string sheetName = null, string dateTime = null, string dateRange = null)`：定義設定報表背景資訊
* `int getColCount()`：定義取得欄位數量
* `void setCol<T>(T[] data)`：定義設定欄位資訊
* `void setCol(DataTable data)`：定義設定欄位資訊


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
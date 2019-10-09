# ReportX
ReportX 可以用簡單的方法，快速建立 Word 與 Excel 報表。更棒的是，它的輸出結果也可以在網頁中直接呈現

## Installation

**Package Manager**

```text
PM> Install-Package ReportX -Version 1.2.0
```

## System requirement

* `v1.2.0` 開始的版本僅支援 .NET Framework 4.5 以上

## API Reference

* ExcelReport
* WordReport
* OdtReport  `2019.09.17 update`
* OdsReport  `2019.09.17 update`
* ReportCreator `2019.09.18 update`

## Default Model

* 以下範例 Model：

```csharp
namespace ReportXTests2.Model
{
    public class ModelEmployeeTicket
    {
        [Present("ID")]
        public Int64 postpid { get; set; }
        [Present("標題")]
        public string posttitle { get; set; }
        [Present("姓名")]
        public string name { get; set; }
        [Present("編號")]
        public string number{ get; set; }
        [Present("資料")]
        public string data { get; set; }
        [Present("電話")]
        public string tel { get; set; }
    }
}
```

## Default

* `v1.2.0` 使用內建規則產生報表，使用範例如下：  

```csharp
//範例: 原始資料
ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
for (int i = 50 - 1; i >= 0; i--)
{
    string s = Guid.NewGuid().ToString("N");
    ModelEmployeeTicket tmp = new ModelEmployeeTicket
    {
        postpid = i+1,
        posttitle = "測試_" + i,
        name = "SOL_" + i,
        number = "123 ",
        data = s,
        tel = "0923456789"
    };
    data[i] = tmp;
}

//範例: 欲顯示哪些標題
string[] cols = new string[5];
    cols[0] = "姓名";
    cols[1] = "資料";
    cols[2] = "ID";
    cols[3] = "電話";

//範例: 標題,產生者
string title = "今日工事";
string Creator = "測試人員";
```
產生標準版報表 使用 ReportCreator<T>
```csharp
//宣告 ReportCreator(以excel為例)
ReportCreator<ExcelReport> report = new ReportCreator<ExcelReport>();
//產生報表字串
string excel = report.render<ExcelReport,ModelEmployeeTicket>(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
//string excel = report.render<ExcelReport>(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
Assert.IsNotNull(excel);
if (File.Exists("creator.xls"))
{
    File.Delete("creator.xls");
    File.AppendAllText("creator.xls", excel);
}
else
{
    File.AppendAllText("creator.xls", excel);
}
```
產生Excel 報表

```csharp
ExcelReport report = new ExcelReport(typeof(ModelEmployeeTicket));
if (cols.Length > 0)
{
    report.setcut(cols);
}
report.setTile(title);
report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
report.setCreator(Creator);
report.setCreatedDate();
report.setColumn();

//data資料型態 可為 model or DataTable
report.setData(data);
report.setsum(data);


//產生報表
var rpData =report.render(null);
Assert.IsNotNull(rpData);
if (File.Exists("report.xls"))
{
    File.Delete("report.xls");
    File.AppendAllText("report.xls", rpData);
}
else
{
    File.AppendAllText("report.xls", rpData);
}

```
產生Word 報表
```csharp
 WordReport report = new WordReport(typeof(ModelEmployeeTicket));
if (cols.Length > 0)
{
    report.setcut(cols);
}
report.setTile(title);
report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
report.setCreator(Creator);
report.setCreatedDate();
report.setColumn();

//data資料型態 可為 model or DataTable
report.setData(data);
report.setsum(data);

//產生報表
var rpData = report.render(null);
Assert.IsNotNull(rpData);
if (File.Exists("report.doc"))
{
    File.Delete("report.doc");
    File.AppendAllText("report.doc", rpData);
}
else
{
    File.AppendAllText("report.doc", rpData);
}

```

`2019/09/17` 新增openOffice(Odt,Ods)   宣告 `Odt`,`Ods`

odt、ods 報表 皆為壓縮檔，皆可解壓縮，解壓縮完會有兩個檔案
* content.xml
* META-INF/manifest.xml

產生Odt 報表
```csharp
 OdtReport report = new OdtReport(typeof(ModelEmployeeTicket));
if (cols.Length > 0)
{
    report.setcut(cols);
}
report.setTile(title);
report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
report.setCreator(Creator);
report.setCreatedDate();
report.setColumn();

//data資料型態 可為 model or DataTable
report.setData(data);
report.setsum(data);

//產生META-INF(OpenOffice設定檔 META-INF/manifest.xml)
report.CreateMeta("odt");
var width = report.getColCount();
//產生報表
var rpData = report.render(width);
Assert.IsNotNull(rpData);
if (File.Exists("content.xml"))
{
    File.Delete("content.xml");
    File.AppendAllText("content.xml", rpData);
}
else
{
    File.AppendAllText("content.xml", rpData);
}
if (File.Exists("content.xml"))
{
    string inputFile = @"content.xml";
    string inputData = @"META-INF/manifest.xml";
    using (var zip = new ZipFile())
    {
        zip.AddFile(inputFile);
        zip.AddFile(inputData);
        zip.Save(@"./report.odt");
    }
}
```
產生Ods 報表
```csharp
OdsReport report = new OdsReport(typeof(ModelEmployeeTicket));
if (cols.Length > 0)
{
    report.setcut(cols);
}
report.setTile(title);
report.setDate(DateTime.Now.AddDays(-1), DateTime.Now);
report.setCreator(Creator);
report.setCreatedDate();
report.setColumn();

//data資料格式 可為 model or DataTable
report.setData(data);
report.setsum(data);

//產生META-INF(OpenOffice設定檔 META-INF/manifest.xml)
report.CreateMeta("ods");

//產生報表
var rpData = report.render();
Assert.IsNotNull(rpData);
if (File.Exists("content.xml"))
{
    File.Delete("content.xml");
    File.AppendAllText("content.xml", rpData);
}
else
{
    File.AppendAllText("content.xml", rpData);
}
if (File.Exists("content.xml"))
{
    string inputFile = @"content.xml";
    string inputData = @"META-INF/manifest.xml";
    using (var zip = new ZipFile())
    {
        zip.AddFile(inputFile);
        zip.AddFile(inputData);
        zip.Save(@"./report.ods");
    }
}
```


## Customized Word and Excel

* `v1.2.0` 自訂表格排序和欄位，可以製作成`Word`和`excel`檔，使用範例如下：

### 自定義欄位

|Funtion_Name      |Content|Type|Example|
|-------------|-------------|-----------|---------|
|setTile    |表格標題|string      |setTile("`表格標題`")|
|setDate    |表格日期|DateTime    |setDate(`starting`, `ending`)|
|setCreator|製表人|string         |setCreator("`作者`")|
|setCreatedDate  |製表時間`DateTime.Now`|`null`  |setCreatedDate()|
|setColumn |表格屬性|`null`    |setColumn()|
|setData   |表格內容  |T[]data/Datatable|setData(data)|
|setcut    |欲顯示欄位|string[] cols| setcut(cols)|
|setsum    |總筆數欄位|T[]data/Datatable|setsum(data)|
----------------------------------------------------------
範例模型
```csharp
ModelEmployeeTicket[] data = new ModelEmployeeTicket[50];
for (int i = 50 - 1; i >= 0; i--)
{
    string s = Guid.NewGuid().ToString("N");
    ModelEmployeeTicket tmp = new ModelEmployeeTicket
    {
        postpid = i + 100,
        posttitle = "測試_" + i,
        name = "SOL_" + i,
        number = "123 ",
        data = s,
        tel = "0923456789"+i
    };
    data[i] = tmp;
}
```
DataTable範例模型

```csharp
DataTable dtTable = new DataTable("dTable");
            DataRow row;
            DataColumn[] colss ={
                    new DataColumn("ID",typeof(int)),
                    new DataColumn("標題",typeof(string)),
                    new DataColumn("姓名",typeof(string)),
                    new DataColumn("編號",typeof(decimal)),
                    new DataColumn("資料",typeof(string)),
                    new DataColumn("電話",typeof(string))
                              };
            dtTable.Columns.AddRange(colss);
            // 建立欄位
            // 新增資料到DataTable
            for (int i = 1; i <= 10; i++)
            {
                    string a = Guid.NewGuid().ToString("N");
                    row = dtTable.NewRow();
                    row["ID"] = i;
                    row["標題"] = "測試 " + i.ToString();
                    row["姓名"] = "SOL_" + i;
                    row["編號"] = "123";
                    row["資料"] = a.ToString();
                    row["電話"] = "0923456789";
                    dtTable.Rows.Add(row);
            }
```

自定義表格欄位
* warning
自定義欄位必須按照：

[架構表格]->[塞入資料]->[加入總筆數] 順序，否則會噴錯!
```csharp
//宣告要產生的檔案類型 方法(以excel為例)
ExcelReport file = new ExcelReport(typeof(ModelEmployeeTicket));

    file.setTile("標題");//標題
    file.setDate(DateTime.Now.AddDays(-1), DateTime.Now);//日期
    file.setCreatedDate();//時間
    file.setColumn();//架構表格
    file.setData(data);//塞入資料
    file.setsum(data);//加入總筆數

    //產生 excel檔
    string excel = file.render(null, "excel");
    File.AppendAllText("Excel檔案.xls", excel);
```

## Preview

* Excel

![excel](https://i.imgur.com/heC8f8i.png)

* Word

![word](https://i.imgur.com/CQCqfcu.png)

* Odt

![Odt](http://192.168.1.136/uploads/-/system/personal_snippet/41/3960d355fd93e84e810a0fee998d2a9a/odtpt.jpg)

* Ods

![Ods](http://192.168.1.136/uploads/-/system/personal_snippet/41/04ac6524c297385718df1633346a1f75/odspt.jpg)

## UML

![UML](http://192.168.1.136/uploads/-/system/personal_snippet/41/6e5d467c63555edd760d163fe5796121/ReportX3__1_.png)

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
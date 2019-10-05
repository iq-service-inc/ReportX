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

* Report
* ExcelReport：
* WordReport
* FileReport  `2018.10.01 update` 
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
    
//範例: 標題
string title = "今日工事";
```

宣告使用Report方法
```csharp
ReportCreator<T> report = new ReportCreator<T>();
```
帶入參數產生Excel 

```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
//ReportCreator<ExcelReport> ex = report.ExcelReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
ReportCreator<ExcelReport> ex = report.ExcelReport(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);

//產生excel 報表
string exce; = ex.render();
if (File.Exists("excel檔案.excel")) File.Delete("excel檔案.excel");
//另存為excel檔
File.AppendAllText("excel檔案.xls", excel); 

```
帶入參數產生Word 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
//ReportCreator<WordReport> wd = report.WordReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
ReportCreator<WordReport> wd = report.WordReport(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
//產生word 報表
string word = wd.render();
if (File.Exists("word檔案.doc")) File.Delete("word檔案.doc");
//另存為word檔
File.AppendAllText("word檔案.doc", word );  

```
`2018/10/01` 新增綜合版   宣告 `FileReport `
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
FileReport file = rep.FileReport(data, cols, title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "SOL", true);
//若要產生 word檔
string word = file.render(null, "word");
//若要產生 excel檔
string excel = file.render(null, "excel");

//另存為Word檔
File.AppendAllText("word檔案.doc", word );
//另存為Excel檔
File.AppendAllText("excel檔案.xls", excel );  
```
`2019/09/17` 新增openOffice(Odt,Ods)   宣告 `Odt `,`Ods `   
帶入參數產生Odt 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
//ReportCreator<OdtReport> odtr = report.OdtReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
ReportCreator<OdtReport> odtr = report.OdtReport(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
//產生odt 報表
string odt = odtr.render();
//產生META-INF(OpenOffice設定檔)
   odtr.CreateMeta("odt");
//壓縮檔案成odt
              if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", odt);
            }
            else
            {
                File.AppendAllText("content.xml", odt);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./odt檔案.odt");
                }
            }

```
帶入參數產生Ods 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
//ReportCreator<OdsReport> odsr = report.OdsReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
ReportCreator<OdsReport> odsr = report.OdsReport(dtData, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "測試人員", true);
//產生ods 報表
string ods = odsr.render();
//產生META-INF(OpenOffice設定檔 META-INF/manifest.xml)
   odsr.CreateMeta("ods");
//壓縮檔案成ods
         if (File.Exists("content.xml"))
            {
                File.Delete("content.xml");
                File.AppendAllText("content.xml", ods);
            }
            else
            {
                File.AppendAllText("content.xml", ods);
            }
            if (File.Exists("content.xml"))
            {
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./ods檔案.ods");
                }
            }
            

```
`2019/09/18` 新增綜合版(包括Odt,Ods)   宣告 `ReportCreator<T> `
帶入參數，使用ReportCreator
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
 ReportCreator<WordReport> wd = res.WordReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
 ReportCreator<ExcelReport> exc = res.ExcelReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
 ReportCreator<OdtReport> orp = res.OdtReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
 ReportCreator<OdsReport> osp = res.OdsReport(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "林家弘", true);
//產生報表
string word = wd.render();
string excel = exc.render();
string ods = odsRes.render(width); (odt要帶入寬度width)
string odsRes = osp.render();
//另存為Word檔
File.AppendAllText("word檔案.doc", word );
//另存為Excel檔
File.AppendAllText("excel檔案.doc", excel );  
//產生META-INF(OpenOffice設定檔)
orp.CreateMeta("odt");
orp.CreateMeta("ods");
//壓縮檔案成odt,ods
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                using (var zip = new ZipFile())
                {
                    zip.AddFile(inputFile);
                    zip.AddFile(inputData);
                    zip.Save(@"./ods檔案.ods");
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
|setCreatedDayRange |報表時間範圍| string |setCreatedDayRange(firstday, lastdday); `2019.09.17 update`
|setColumn |表格屬性|`null`    |setColumn()|
|setData   |表格內容  |T []data     |setData(data)|
|setcut    |欲顯示欄位|string[] cols| setcut(cols)|
|setsum    |總筆數欄位|T []data|setsum(data)|


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
//宣告FileReport 方法
FileReport file = new FileReport(typeof(ModelEmployeeTicket));

    file.setTile("標題");//標題
    file.setDate(DateTime.Now.AddDays(-1), DateTime.Now);//日期
    file.setCreatedDate();//時間
    file.setColumn();//架構表格
    file.setData(data);//塞入資料
    file.setsum(data);//加入總筆數

    //產生 word檔
    string word = file.render(null, "word");
    File.AppendAllText("自定義綜合版.doc", word);

    //產生 excel檔
    string excel = file.render(null, "excel");
    File.AppendAllText("自定義綜合版.xls", excel);
```


## Multi ExcelWorksheet
  在Excel 做分頁表格

### 參數
ExcelReport 的 陣列

```csharp
List<ExcelReport> excelResList = new List<ExcelReport>();
MultiExcel multiExcel = new MultiExcel(excelResList);
string res = multiExcel.render();
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
![UML](http://192.168.1.136/uploads/-/system/personal_snippet/41/bb07e5c250d13d472e18f4bd53d96734/ReportX2.png)


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
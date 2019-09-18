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
* FileReport  `2018.10.01 updata` 
* OdtReport  `2019.09.17 updata` 
* OdsReport  `2019.09.17 update` 
* ReportCreatorReport `2019.09.18 update`


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
Report Rpt = new Report();
```
帶入參數產生Excel 

```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
ExcelReport er = Rpt.excelResponse(data , cols, title , DateTime.Now.AddDays(-1), DateTime.Now, "SOL", true);

//產生excel 報表
string exce; = er.render();
if (File.Exists("excel檔案.doc")) File.Delete("excel檔案.doc");
//另存為excel檔
File.AppendAllText("excel檔案.xls", excel); 

```
帶入參數產生Word 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
WordReport wr =Rpt.WordResponse(data, cols , title, Convert.ToDateTime("2017-01-20"), Convert.ToDateTime("2017-01-20"), "SOL",true);
//產生word 報表
string word = wr.render();
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
File.AppendAllText("excel檔案.doc", excel );  
```
`2019/09/17` 新增openOffice(Odt,Ods)   宣告 `Odt `,`Ods `   
帶入參數產生Odt 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
    Odt odtRes = Rpt.OdtResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "SOL", true);
//產生odt 報表
string odt = odtRes.render();
//產生META-INF(OpenOffice設定檔)
   odtRes.CreateMeta("odt");
//壓縮檔案成odt
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @".\result.odt";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {
                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            

```
帶入參數產生Ods 報表
```csharp
//報表 (原始資料 ,欄位陣列 , 標題 , 開始時間 , 結束時間 , 製表人 ,是否顯示結尾(總筆數)欄位)
    Ods odsRes = Rpt.OdtResponse(data, cols, title, DateTime.Now.AddDays(-1), DateTime.Now, "SOL", true);
//產生ods 報表
string ods = odsRes.render();
//產生META-INF(OpenOffice設定檔)
   odtRes.CreateMeta("ods");
//壓縮檔案成ods
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @".\result.ods";
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {
                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
            

```
`2019/09/18` 新增綜合版(包括Odt,Ods)   宣告 `ReportCreator<T> `
帶入參數，使用ReportCreator
```csharp=
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
                string[] input = new string[2];
                string inputFile = @"content.xml";
                string inputData = @"META-INF/manifest.xml";
                input[0] = inputFile;
                input[1] = inputData;
                string outputFile = @".\result.ods";(副檔名要改)
                using (var output = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                    try
                    {
                        using (var zip = new ZipOutputStream(output))
                        {
                            zip.SetLevel(9);
                            byte[] buffer = new byte[4096];
                            foreach (string file in input)
                            {
                                ZipEntry entry = new ZipEntry(file);
                                entry.DateTime = DateTime.Now;
                                zip.PutNextEntry(entry);
                                using (FileStream fs = System.IO.File.OpenRead(file))
                                {
                                    int sourceBytes;
                                    do
                                    {
                                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                        zip.Write(buffer, 0, sourceBytes);
                                    } while (sourceBytes > 0);
                                }
                            }
                            zip.Finish();
                            zip.Close();
                        }
                    }
                    catch (Exception ex)
                    {
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
![Odt](https://i.imgur.com/ENBBLp2.jpg)
* Ods
![Ods](https://i.imgur.com/9Ij8V8q.jpg)
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
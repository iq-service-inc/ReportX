# ReportX
ReportX是一個快速建立Word.doc 和 Excel.xls 報表之API，幫助開發人員能夠快速完成報表輸出之功能。

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
  * 可以自訂報表規則，使用範例如下：  

```csharp

ExcelReport Exx = new ExcelReport(typeof(Modeltrend)); //使用ExcelReport 方法

Exx.setTile(title);  //設置標題
Exx.setDate(Convert.ToDateTime(starttime), Convert.ToDateTime(endtime)); //自訂時間區間
Exx.setCreatedDate();  //製表時間
Exx.setColumn(); //建立表格屬性
Exx.setData(data); //匯入資料 (Model)
            
//統計資料數
Exx.appendRow(new { value = "總筆數", colspan = Exx.getColCount() - 1, style = lastRowStyle }, data.Length);
            
//輸出報表
string output = Exx.render();
            
```

## Default

* `v1.2.0`
  * 可以使用內建報表規則，使用範例如下：  

```csharp

Report s = new Report(); //使用Report 方法

//帶入參數(資料,開始時間,結束時間,製表人)
ExcelReport myca = s.excelResponse(data,"Report", Convert.ToDateTime(starttime), Convert.ToDateTime(endtime), "SOL");
//統計資料數
myca.appendRow(new { value = "筆數", colspan = myca.getColCount() - 1, style = lastRowStyle }, data.Length);
            
//輸出報表
string output = Exx.render();
            
```
## Output 

 * 輸出報表需轉存成Word或Excel ，請建立以下報表輸出方法：  

```csharp

public HttpResponseMessage getFile(string content, string fileNmae = null, string mimeType = "application/octet-stream", HttpStatusCode code = HttpStatusCode.OK)
{
    HttpResponseMessage resp = new HttpResponseMessage();
    resp.StatusCode = code;
    string fn = (fileNmae ?? Guid.NewGuid().ToString()) + ".xls";//根據需求更改.xls 或 .doc
    if (content != null)
     {
       resp.Content = new StringContent(content, Encoding.UTF8, mimeType);
       resp.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
        {
         FileName = fn
        };
      }
    return resp;
}

```
 * 輸出報表：  

```csharp

    return getFile(output);
    
```

* 若嫌麻煩,，請直接使用Zaplib的 getAttachmentResponse()方法！
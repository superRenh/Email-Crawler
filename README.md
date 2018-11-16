# Email Crawler on POP server
連接pop伺服器自動地爬取郵件，將Email檔案轉變為結構化的資料並儲存於MSSQL資料庫，以利後續的應用。
<ul>
  <li>爬取目標：</li>
  給定Email帳號密碼、POP伺服器名稱、連接埠，以POP3協定運作之所有Email。此處以爬取個人OUTLOOK信箱為例，需於OUTLOOK中設定權限授權：
  <img src="https://github.com/superRenh/Email-Crawler/blob/master/images/pop%E8%A8%AD%E5%AE%9A.JPG" width="80%" height="80%" style="float.center">
  <li>爬取欄位：</li>
  Email的標題、內容、寄件者、收件者、寄件時間、Message ID、附件檔案(另存於指定資料夾)、詞頻(Term Frequency) 
  <li>Database資料表：</li>
    <ul>
      <li>eml_log：</li>
  <table style="width:100%">
  <tr>
    <th>欄位名稱</th>
    <th>欄位說明</th> 
    <th>資料格式</th>
  </tr>
  <tr>
    <td>Eml_datetime</td>
    <td>郵件時間</td> 
    <td>datetime</td>
  </tr>
  <tr>
    <td>type</td>
    <td>掃描狀態</td> 
    <td>nvarchar(50)</td>
  </tr>
</table>
  </ul>
</ul>

## <ins>Environment setup環境設定(Python3.6, MSSQL)<ins>

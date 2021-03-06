# Email Crawler on POP server
程式開發採用Python 3.6及MSSQL 2012版本。連接pop伺服器自動地爬取郵件，依特定掃描規則(主旨、寄件者等)爬取Email。使用正則表達式解析內文並將分析的結果存入資料庫，並下載Email附件檔案到指定路徑。最後於pop伺服器端刪除已存入資料庫的郵件。此專案的目標為將Email檔案，轉變為結構化的資料並儲存於資料庫，以利後續的應用。
<ul>
  <li>目標：</li>
  給定Email帳號密碼、POP伺服器名稱、連接埠，以POP3協定運作之所有Email。此處以爬取個人OUTLOOK信箱為例，需於OUTLOOK中設定權限授權：
  <img src="https://github.com/superRenh/Email-Crawler/blob/master/images/pop%E8%A8%AD%E5%AE%9A.JPG" width="80%" height="80%" style="float.center">
  <li>爬取欄位：</li>
  Email的標題、內容、寄件者、收件者、寄件時間、Message ID、附件檔案(另存於指定資料夾)、並從內容中分析出每封信的詞頻(Term Frequency)
  <li>運行模式：</li>
    <ol type="I">
      <li>將爬取過的Email存入eml_log資料表，後續用Uid判別不重複爬取。</li>
      <li>依爬取規則將符合目標之Email儲存於email_pop資料表，並將附件檔案另存於指定資料夾。</li>
      <li>將存入email_pop資料表的Email，於POP Server端刪除。</li>
    </ol>
  <li>爬取規則：</li>
    <ol type="1">
      <li>排除已爬取過並存入資料庫的Email，僅爬取新增的Email，縮短每次爬取的時間。</li>
      <li>指定寄件者格式，或不指定全部爬取。</li>
      <li>可以選擇排除沒有附件檔案的email。</li>
    </ol>
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
    <td>ID</td>
    <td>table index</td> 
    <td>int</td>
  </tr>
  <tr>
    <td>Eml_datetime</td>
    <td>寄件時間</td> 
    <td>datetime</td>
  </tr>
  <tr>
    <td>type</td>
    <td>掃描狀態</td> 
    <td>nvarchar(50)</td>
  </tr>
  <tr>
    <td>time</td>
    <td>執行時間</td> 
    <td>datetime</td>
  </tr>
  <tr>
    <td>MessageID</td>
    <td>Email ID</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Subject</td>
    <td>標題</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Uid</td>
    <td>POP ID</td> 
    <td>nvarchar(MAX)</td>
  </tr>
</table>
      <li>email_pop：</li>
      <table style="width:100%">
  <tr>
    <th>欄位名稱</th>
    <th>欄位說明</th> 
    <th>資料格式</th>
  </tr>
  <tr>
    <td>ID</td>
    <td>table index</td> 
    <td>int</td>
  </tr>
  <tr>
    <td>Run_time</td>
    <td>執行時間</td> 
    <td>datetime(50)</td>
  </tr>
  <tr>
    <td>Eml_datetime</td>
    <td>寄件時間</td> 
    <td>datetime</td>
  </tr>
  <tr>
    <td>Eml_from</td>
    <td>寄件者</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Eml_to</td>
    <td>收件者</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Subject</td>
    <td>標題</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Content</td>
    <td>內容</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Pdf_name</td>
    <td>附檔名稱</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>MessageID</td>
    <td>Email ID</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>Uid</td>
    <td>POP ID</td> 
    <td>nvarchar(MAX)</td>
  </tr>
  <tr>
    <td>TFs</td>
    <td>Email內容詞頻分析</td> 
    <td>nvarchar(MAX)</td>
  </tr>
</table>
  </ul>
</ul>

## <ins>環境設定(Python3.6, MSSQL)<ins>
1. 下載資料庫套件</br>
`pip install pyodbc`
2. 下載email套件</br>
`pip install email`</br>
3. 下載文本分析套件</br>
`pip install jieba`</br>
4. 下載連接pop server套件</br>
`pip install poplib`</br>
## <ins>config檔設定<ins>
<ol>
  <li>"DB": 設定資料庫相關訊息driver, server, database name, uid, password。</li>
  <li>"email_from":設定指定爬取的寄件者格式，或設置為空，不做限制。</li>
  <li>"save": 設定附件檔案下載路徑。</li>
</ol>
<img src="https://github.com/superRenh/Email-Crawler/blob/master/images/config.JPG" width="80%" height="80%" style="float.center">

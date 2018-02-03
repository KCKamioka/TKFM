#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ 画面モック自動ログイン																						Ver.1.0
#_/ 
#_/ 【引数】
#_/ Args[0]								Args[1]			Args[2]				Args[3]		Args[4]			Args[5]
#_/ 接続サーバー						システム		事務				端末		リフレッシュ	パーソナルエリア
#_/ --------------------------------------------------------------------------------------------------------------------
#_/ 1：障害者福祉　京都市環境			9：障害者福祉	0：障害者手帳		 0：PC0010	0：Off			0：Off
#_/ 2：障害者福祉　開発環境(114環境)					1：補装具			 1：PC0012	1：On			1：On
#_/ 3：障害者福祉　レビュ環境(122環境)					2：障害者総合支援	 2：PC0014
#_/ 4：国保　京都市環境									3：更生医療			 3：PC0016
#_/ 5：国保　開発環境									4：障害者福祉・共通	 4：PC0020
#_/ 6：国保　レビュ環境									5：特別障害者手当	 5：PC0022
#_/ 7：介護　京都市環境									6：心身扶養共済		 6：PC0024
#_/ 8：介護　開発環境									7：所属長			 7：PC0026
#_/ 9：介護　レビュ環境									8：管理者			 8：PC0028
#_/ 																		 9：PC0030
#_/ 																		10：PC0032
#_/ 																		11：PC0034
#_/ 																		12：PC0036
#_/ 																		13：PC0038
#_/ 																		14：PC0040
#_/ 																		15：PC0099
#_/ 
#_/ 																										SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing
#ハッシュ
$settings = @{}
#インポートした項目をハッシュ変数にマッピングしながら格納
Import-Csv -Path ..\Config\setting.ini -Header Key,Value -Delimiter "=" | %{$settings.Add($_.Key.Trim(),$_.Value.Trim())}

#ＵＲＬ
switch($Args[0]){
  1 {$url="https://"+$settings.SYOGAI_KYOTO_SERVER+$settings.URL_VALUE1}		#syogai-dc.open.kyoto.jp
  2 {$url="https://"+$settings.SYOGAI_DEVELOP_SERVER+$settings.URL_VALUE1}		#210.134.125.114
  3 {$url="https://"+$settings.SYOGAI_REVIEW_SERVER+$settings.URL_VALUE1}		#210.134.125.122
  4 {$url="https://"+$settings.KOKUHO_KYOTO_SERVER+$settings.URL_VALUE1}
  5 {$url="https://"+$settings.KOKUHO_DEVELOP_SERVER+$settings.URL_VALUE1}
  6 {$url="https://"+$settings.KOKUHO_REVIEW_SERVER+$settings.URL_VALUE1}
  7 {$url="https://"+$settings.KAIGO_KYOTO_SERVER+$settings.URL_VALUE1}
  8 {$url="https://"+$settings.KAIGO_DEVELOP_SERVER+$settings.URL_VALUE1}
  9 {$url="https://"+$settings.KAIGO_REVIEW_SERVER+$settings.URL_VALUE1}
  default {Write-Output '変数$Args[0]は1～9ではありません。'}
}

if($Args[5] -eq "1") {
	#パーソナルエリアの設定
	$url=$url+$settings.PERSONAL_AREA
}

#画面入力値
$strUser		= $settings.MOCK_USER
$strPass		= $settings.MOCK_PASS
$strSystem		= $Args[1]
$strBusiness	= $Args[2]
$strTerminalId	= $Args[3]

#IE起動
$ie = New-Object -ComObject InternetExplorer.Application

# IE画面調整ここから
$ie.top			= $settings.WINDOW_TOP
$ie.left		= $settings.WINDOW_LEFT
$ie.width		= 600
$ie.height		= 400
$ie.StatusBar	= $false
$ie.ToolBar		= $false
$ie.MenuBar		= $false
$ie.AddressBar	= $false
$ie.Resizable	= $false

# BLANKページを開いて，ウィンドウ内サイズを取得する
$ie.Navigate("about:blank")
$ie.Document.IHTMLDocument2_write(@"
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja" lang="ja">
<meta http-equiv="Content-Type" content="text/html;charset=Shift_JIS">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<title>JavaScript テスト</title>
<script type="text/javascript">
getWindowSize();

function getWindowSize() {
	sW3 = parseInt(window.innerWidth);
	sH3 = parseInt(window.innerHeight);
	document.write(sW3+":"+sH3);
}
</script>
</head>
<body>

</body>
</html>
"@)

# ページをロードし終わるまで待つ
while ($ie.Busy) {sleep -milliseconds 500}

# ウィンドウ内サイズを1280*1002にする
$test = $ie.Document.IHTMLDocument2_body.innerText
$size = $test.Split(":")
$ie.width=600+(1280-$size[0])
$ie.height=400+(1002-$size[1])

# 目的のURLを開く
$ie.Navigate($url)

# ページをロードし終わるまで待つ
while ($ie.Busy) {sleep -milliseconds 500}

# IEを可視化する
$ie.Visible = $true

# IEをアクティブにする
add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
$b = (Get-Process | Where-Object {$_.MainWindowTitle -like "*Internet Explorer*"}).id
start-sleep -Milliseconds 500
[Microsoft.VisualBasic.Interaction]::AppActivate($b)

# 証明書での続行
#$element = [System.__ComObject].InvokeMember(
#			"getElementById"										# 指定方法にIDを使用
#			,[System.Reflection.BindingFlags]::InvokeMethod			# おまじない
#			, $null													# 不要パラメータ
#			, $ie.Document											# ページ情報オブジェクト
#			, "overridelink"										# ID名
#			)
#if($element -ne [System.DBNull]::Value) {
#	$element.click()
#}
#
# ページをロードし終わるまで待つ
#while ($ie.Busy) {sleep -milliseconds 500}

# ログイン情報設定
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtUsernameInput").value=$strUser
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtPasswordInput").value=$strPass
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
start-sleep -milliseconds 1000
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtBusinessIdComboBox").SelectedIndex=$strBusiness
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtTerminalIdComboBox").SelectedIndex=$strTerminalId
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtLoginButton").click()

# ページをロードし終わるまで待つ
while ($ie.Busy) {sleep -milliseconds 1000}
start-sleep -Milliseconds 1000

if($Args[4] -eq "1") {
	start-sleep -milliseconds 3000
	#リフレッシュ起動
	start-process powershell.exe -windowstyle hidden .\Mock_reflesh.ps1
	#モニタリング起動
	start-process powershell.exe -windowstyle hidden .\Mock_monitoring.ps1
}

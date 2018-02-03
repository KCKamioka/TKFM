#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ 画面モック自動ログイン																						Ver.1.0
#_/ 
#_/ 【引数】
#_/ Args[0]							Args[1]				Args[2]				Args[3]		Args[4]			Args[5]
#_/ 接続サーバー					システム			事務				端末		リフレッシュ	パーソナルエリア
#_/ --------------------------------------------------------------------------------------------------------------------
#_/ 1：住基　　　開発環境			10：福祉医療		0：老人医療			 0：PC0010	0：Off			0：Off
#_/ 2：住基　　　統合開発環境		11：敬老乗車証		1：障害者医療		 1：PC0012	1：On			1：On
#_/ 3：福祉医療　開発環境			12：住基・税照会	2：ひとり親医療		 2：PC0014
#_/ 4：福祉医療　統合開発環境		13：共通			3：子ども医療		 3：PC0016
#_/ 													4：所属長			 4：PC0020
#_/ 													5：管理者			 5：PC0022
#_/ 																		 6：PC0024
#_/ 																		 7：PC0026
#_/ 																		 8：PC0028
#_/																			 9：PC0030
#_/ 																		10：PC0032
#_/ 																		11：PC0034
#_/ 																		12：PC0036
#_/ 																		13：PC0038
#_/ 																		14：PC0040
#_/ 																		15：PC0088
#_/ 																		16：PC0099
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
  1 {$url="http://"+$settings.JUKI_DEVELOP_SERVER+$settings.URL_VALUE1}			#192.168.60.24
  2 {$url="https://"+$settings.JUKI_REVIEW_SERVER+$settings.URL_VALUE1}			#juki-dc.open.kyoto.jp
  3 {$url="http://"+$settings.FUKUSHI_DEVELOP_SERVER+$settings.URL_VALUE1}		#192.168.60.21
  4 {$url="https://"+$settings.FUKUSHI_REVIEW_SERVER+$settings.URL_VALUE1}		#fukui-dc.open.kyoto.jp
  default {Write-Output '変数$Args[0]は1～4ではありません。'}
}

if($Args[5] -eq "1") {
	#パーソナルエリアの設定
	$url=$url+$settings.PERSONAL_AREA
}

#画面入力値

# 「福祉医療（統合開発）」の場合ユーザー・パスは固定
if($Args[0] -eq "4") {
    $strUser		= "admin"
    $strPass		= "admin"
}else{
    $strUser		= $settings.MOCK_USER
    $strPass		= $settings.MOCK_PASS
}

$strSystem		= $Args[1]
$strBusiness	= $Args[2]
$strTerminalId	= $Args[3]

#IE起動
$ie = New-Object -ComObject InternetExplorer.Application

# IE画面調整ここから
$ie.top			= $settings.WINDOW_TOP
$ie.left		= $settings.WINDOW_LEFT
$ie.width		= 300
$ie.height		= 200
$ie.StatusBar	= $false
$ie.ToolBar		= $false
$ie.MenuBar		= $false
$ie.AddressBar	= $false
$ie.Resizable	= $false
$ie.Visible 	= $false

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
$ie.width=300+(1280-$size[0])
$ie.height=200+(1002-$size[1])

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

# ログイン情報設定
# 「住基」の場合ユーザー・パスは設定しない
if($Args[1] -ge "3") {
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtUsernameInput").value=$strUser
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtPasswordInput").value=$strPass
}

# # 「福祉医療（開発環境）」の場合ボタン指定
# if($Args[0] -eq "3") {
#     #システム
#     switch($Args[1]){
#         10 {
#             #事務
#             switch($Args[2]){
#                 0 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt60").click()}		#老人医療
#                 1 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt20").click()}		#障害者医療
#                 2 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt43").click()}		#ひとり親医療
#                 3 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt56").click()}		#子ども医療
#                 default {Write-Output '変数$Args[0]は1～4ではありません。'}
#             }
#         }		#福祉医療
#         11 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt118").click()}		#敬老乗車証
#         12 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt46").click()}		#住基・税照会
#         13 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
#             sleep -milliseconds 500
#             [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
#             }		#共通
#         default {Write-Output '変数$Args[0]は1～4ではありません。'}
#     }
#     sleep -milliseconds 1000
# }else{
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
    sleep -milliseconds 500
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
    sleep -milliseconds 500
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtBusinessIdComboBox").SelectedIndex=$strBusiness
# }

# 端末
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtTerminalIdComboBox").SelectedIndex=$strTerminalId
# ログイン
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtLoginButton").click()

# ページをロードし終わるまで待つ
while ($ie.Busy) {sleep -milliseconds 3000}

if($Args[4] -eq "1") {
	start-sleep -milliseconds 3000
	#リフレッシュ起動
	start-process powershell.exe -windowstyle hidden .\Mock_reflesh.ps1
	#モニタリング起動
	start-process powershell.exe -windowstyle hidden .\Mock_monitoring.ps1
}

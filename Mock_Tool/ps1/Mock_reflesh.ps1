#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ りふれっしゅ									Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing

#実行プログラム
$execPath = '.\Mock_monitoring.lnk'
#実行プログラム起動
start-process $execPath

# エラー発生で終了
$ErrorActionPreference = "stop"

# シェルオブジェクト生成
$shell = New-Object -ComObject Shell.Application
# IEリスト取得
$ie_list = @($shell.Windows() | where { $_.Name -match "Internet Explorer" })
# 複数ヒットした場合は一番最後を取得する(添え字-1は最後）
$ie = @($ie_list | where { $_.LocationURL -match $V_IPADDR })[-1]

# IEが見つからなくなるとエラーになるのでShellは終了する
do {
	#「Start-Sleep -s 300」で300秒(5分)スリープ
	Start-Sleep -s 300
	# IEをリフレッシュ
	$ie.Refresh()
}while ($true)

exit

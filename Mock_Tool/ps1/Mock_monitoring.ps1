#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ もにたりんぐ									Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing

#監視プロセス名
$procName = "iexplore"

#終了プロセス名
$EndProc = "powershell"

#監視&自動起動
try{
	#プロセスを取得
	$p_lis = Get-Process $procName -ErrorAction SilentlyContinue
	# 複数ヒットした場合は一番最後を取得する(添え字-1は最後）
	$p = @($p_lis | where { $_.LocationURL -match $V_IPADDR })[-1]

	#開始したプロセスが終了するまで待機するように指示
	$p.WaitForExit()
}
catch{
	Write-Host "No Process"
}
finally{
	#プロセス終了
	Stop-process -name $EndProc
}

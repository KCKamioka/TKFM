#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ショートカット作成								Ver.1.0
#_/ 
#_/ 											SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
try{
	# スクリプトが存在するディレクトリを取得
	$base = split-Path -Parent $MyInvocation.MyCommand.Path;

	#シェルオブジェクト作成
	$ws = New-Object -com WScript.Shell;

	# 既存フォルダがある場合、既存のフォルダは何も変わらない
	New-Item "..\MockShortcut" -type directory -Force 

	# 「Mock_Tool.csv」を読み込み、ショートカットを作成する
	foreach ($line in (get-content "$base`\..\Config\Mock_Tool.csv")){
		$s = $line.split(",");
		#ショートカットオブジェクト
		$lnk = $ws.CreateShortcut($base + "\..\MockShortcut\" + $s[0] + ".lnk");
		#リンク先(対象となるファイル)
		$lnk.TargetPath = "powershell";
		#コマンドライン引数
		$lnk.Arguments = $base+ "\" + $s[1] + " " + $s[2];
		#作業フォルダ
		$lnk.WorkingDirectory = $base;
		#アイコン(パス+インデックス)
		$lnk.IconLocation = $s[5] + "," + $s[6];
		#ショートカットキー( "Alt+", "Ctrl+", "Shift+" などと共に)
		$lnk.Hotkey = $s[3];
		#実行時の大きさ(1:通常ウインドウ、3:最大化、7:最小化)
		$lnk.WindowStyle = "7";
		#コメント(ショートカットの説明)
		$lnk.Description = $s[4];
		#プロパティの保存
		$lnk.save();
	};

}finally{
	if ($lnk -ne $null){ 
		while([System.Runtime.InteropServices.Marshal]::ReleaseComObject($lnk) -gt 0){}
	};
	if ($ws -ne $null){ 
		while([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) -gt 0){}
	};
}

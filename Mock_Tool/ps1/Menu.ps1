$nObjSeq=0

#カレントディレクトリ取得
$CurDir=Split-Path $myInvocation.MyCommand.Path -Parent

#ファイルパス取得
$tmp_path = (Convert-Path ..)
$csvFilePth = $tmp_path + "\Config\Mock_Tool.csv"

#設定ファイル存在チェック
if(!(Test-Path($csvFilePth))){
    Read-Host "設定ファイルが見つかりませんでした。"$csvFilePth
    exit 1
}

#csvファイル読込
$param=$null
$line=$null
$sCsvKeys=$null
$sCsvAppName=$null
$sCsvVals=$null
$sCsvOther=$null
foreach ($line in Get-Content $csvFilePth) {
    $param=$line.split(",",4)
    $sCsvKeys +=,$param[0]
    $sCsvAppName +=,$param[1]
    $sCsvVals +=,$param[2]
    $sCsvOther +=,$param[3]
}

#Mock_Login_ArgumentSpecialの設定値を格納
$LoginSettings=$null
$LoginSettings=$sCsvVals[[Array]::IndexOf($sCsvKeys,'Mock_Login_ArgumentSpecial')]
$LoginSettings=$LoginSettings.split(" ")

# アセンブリの読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function CreateShortcut($sArgSpecialLine,$nWindowStyleIdx){

    try{

        #ファイル出力先
        $shortcut=[Environment]::GetFolderPath('Programs')+"\MockShortcut"

	    #シェルオブジェクト作成
	    $ws = New-Object -com WScript.Shell;

	    # 既存フォルダがある場合、既存のフォルダは何も変わらない
	    New-Item $shortcut -type directory -Force 

	    # 「Mock_Tool.csv」を読み込み、ショートカットを作成する
		$s = $sArgSpecialLine.split(",");
		#ショートカットオブジェクト
		$lnk = $ws.CreateShortcut($shortcut+"\" + $s[0] + ".lnk");
		#リンク先(対象となるファイル)
		$lnk.TargetPath = "powershell";
		#コマンドライン引数
		$lnk.Arguments = $CurDir+ "\" + $s[1] + " " + $s[2];
		#作業フォルダ
		$lnk.WorkingDirectory = $CurDir;
		#アイコン(パス+インデックス)
		$lnk.IconLocation = $s[5] + "," + $s[6];
		#ショートカットキー( "Alt+", "Ctrl+", "Shift+" などと共に)
		$lnk.Hotkey = $s[3];
		#実行時の大きさ(1:通常ウインドウ、3:最大化、7:最小化)
		$lnk.WindowStyle = $nWindowStyleIdx;
		#コメント(ショートカットの説明)
		$lnk.Description = $s[4];
		#プロパティの保存
		$lnk.save();

    }finally{
	    if ($lnk -ne $null){ 
		    while([System.Runtime.InteropServices.Marshal]::ReleaseComObject($lnk) -gt 0){}
	    };
	    if ($ws -ne $null){ 
		    while([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) -gt 0){}
	    };
    }
}

try {

    #ラベルとテキストボックスの幅
    $nLabelTextWidth=20

    # フォームの高さ
    $nFormHight=480

    # オブジェクトの間隔
    $nObjSpace=7

    # フォームの作成
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Setting"
    $form.Size = New-Object System.Drawing.Size(260,$nFormHight) 

    # メッセージラベルの設定
    $lblTips = New-Object System.Windows.Forms.Label
    $lblTips.Location = New-Object System.Drawing.Point(10,10) 
    $lblTips.Size = New-Object System.Drawing.Size(200,40) 
    $lblTips.ForeColor = "blue"
    $lblTips.Text = "◇ショートカットキー「Ctrl+Shift+Alt+A」でも起動できます。"

    # 接続サーバーラベルの設定
    $nObjSeq=$nObjSeq+1
    $objHight=$nFormHight*($nObjSeq/$nObjSpace)
    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Location = New-Object System.Drawing.Point(10,$objHight) 
    $lblServer.Size = New-Object System.Drawing.Size(120,20) 
    $lblServer.Text = "接続サーバー："
  
    # 接続サーバーコンボボックスを作成
    $objHight=$objHight+$nLabelTextWidth
    $cmbServer = New-Object System.Windows.Forms.Combobox
    $cmbServer.Location = New-Object System.Drawing.Point(20,$objHight)
    $cmbServer.size = New-Object System.Drawing.Size(200,50)
    $cmbServer.DropDownStyle="DropDownList"

    # 接続サーバーコンボボックスに項目を追加
    $ServerNo = @("3","4")
    $ServerName = @("開発環境","統合開発環境")
    [void] $cmbServer.Items.AddRange($ServerName)

    # 接続サーバーコンボボックスに項目を表示
    $cmbServer.Text=$ServerName[[array]::IndexOf($ServerNo,$LoginSettings[0])]

    # システムラベルの設定
    $nObjSeq=$nObjSeq+1
    $objHight=$nFormHight*($nObjSeq/$nObjSpace)
    $lblSystem = New-Object System.Windows.Forms.Label
    $lblSystem.Location = New-Object System.Drawing.Point(10,$objHight) 
    $lblSystem.Size = New-Object System.Drawing.Size(120,20) 
    $lblSystem.Text = "システム："
  
    # システムコンボボックスを作成
    $objHight=$objHight+$nLabelTextWidth
    $cmbSystem = New-Object System.Windows.Forms.Combobox
    $cmbSystem.Location = New-Object System.Drawing.Point(20,$objHight)
    $cmbSystem.size = New-Object System.Drawing.Size(200,50)
    $cmbSystem.DropDownStyle="DropDownList"

    # システムコンボボックスに項目を追加
    $SystemNo = @("10","11","13","14")
    $SystemName = @("福祉医療","敬老乗車証","住基・税照会（福祉）","共通")
    [void] $cmbSystem.Items.AddRange($SystemName)

    # システムコンボボックスに項目を表示
    $cmbSystem.Text=$SystemName[[array]::IndexOf($SystemNo,$LoginSettings[1])]
    
    # 事務ラベルの設定
    $nObjSeq=$nObjSeq+1
    $objHight=$nFormHight*($nObjSeq/$nObjSpace)
    $lblZimu = New-Object System.Windows.Forms.Label
    $lblZimu.Location = New-Object System.Drawing.Point(10,$objHight) 
    $lblZimu.Size = New-Object System.Drawing.Size(120,20) 
    $lblZimu.Text = "事務："
  
    # 事務コンボボックスを作成
    $objHight=$objHight+$nLabelTextWidth
    $cmbZimu = New-Object System.Windows.Forms.Combobox
    $cmbZimu.Location = New-Object System.Drawing.Point(20,$objHight)
    $cmbZimu.size = New-Object System.Drawing.Size(200,50)
    $cmbZimu.DropDownStyle="DropDownList"

    # 事務コンボボックスに項目を追加
    $ZimuNo = @("0","1","2","3","4","5")
    $ZimuName = @("老人医療","障害者医療","ひとり親医療","子ども医療","所属長","管理者")
    [void] $cmbZimu.Items.AddRange($ZimuName)

    # 事務コンボボックスに項目を表示
    $cmbZimu.Text=$ZimuName[[array]::IndexOf($ZimuNo,$LoginSettings[2])]

    # 端末ラベルの設定
    $nObjSeq=$nObjSeq+1
    $objHight=$nFormHight*($nObjSeq/$nObjSpace)
    $lblKu = New-Object System.Windows.Forms.Label
    $lblKu.Location = New-Object System.Drawing.Point(10,$objHight) 
    $lblKu.Size = New-Object System.Drawing.Size(120,20) 
    $lblKu.Text = "端末："
  
    # 端末コンボボックスを作成
    $objHight=$objHight+$nLabelTextWidth
    $cmbKu = New-Object System.Windows.Forms.Combobox
    $cmbKu.Location = New-Object System.Drawing.Point(20,$objHight)
    $cmbKu.size = New-Object System.Drawing.Size(200,50)
    $cmbKu.DropDownStyle="DropDownList"

    # 端末コンボボックスに項目を追加
    $KuNo = @("0","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16")
    $KuName = @("北区",
                "上京区",
                "左京区",
                "中京区",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "")
    [void] $cmbKu.Items.AddRange($KuName)

    # 端末コンボボックスに項目を表示
    $cmbKu.Text=$KuName[[array]::IndexOf($KuNo,$LoginSettings[3])]

    # OKボタンの設定
    $nObjSeq=$nObjSeq+1
    $objHight=$nFormHight*($nObjSeq/$nObjSpace)+20
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(40,$objHight)
    $btnOK.Size = New-Object System.Drawing.Size(75,30)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = "OK"

    # キャンセルボタンの設定
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(130,$objHight)
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = "Cancel"

    # キーとボタンの関係
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    # ボタン等をフォームに追加
    $form.Controls.Add($lblTips)
    $form.Controls.Add($lblServer) 
    $form.Controls.Add($cmbServer)
    $form.Controls.Add($lblSystem)
    $form.Controls.Add($cmbSystem)
    $form.Controls.Add($lblZimu)
    $form.Controls.Add($cmbZimu)
    $form.Controls.Add($lblKu)
    $form.Controls.Add($cmbKu)
    $form.Controls.Add($btnOK)
    $form.Controls.Add($btnCancel)

    # フォームを表示させ、その結果を受け取る
    $result = $form.ShowDialog()

    # 結果による処理分岐
    if ($result -eq "OK"){

        #csv設定内容を取得
        $LoginSettings[0]=$ServerNo[[array]::IndexOf($ServerName,$cmbServer.Text)]
        $LoginSettings[1]=$SystemNo[[array]::IndexOf($SystemName,$cmbSystem.Text)]
        $LoginSettings[2]=$ZimuNo[[array]::IndexOf($ZimuName,$cmbZimu.Text)]
        $LoginSettings[3]=$KuNo[[array]::IndexOf($KuName,$cmbKu.Text)]
        
        #csvファイル書込
        $sTmpLoginSettings=$null
        $sTmpLoginSettings=[System.String]::Join(" ", $LoginSettings)
        $nArgSpecialIdx=[Array]::IndexOf($sCsvKeys,'Mock_Login_ArgumentSpecial')
        $sCsvVals[$nArgSpecialIdx]=$sTmpLoginSettings
        $line=$null
        $tmpline=$null
        $i=0
        foreach ($key in $sCsvKeys) {
            $tmpline=$sCsvKeys[$i]+","+$sCsvAppName[$i]+","+$sCsvVals[$i]+","+$sCsvOther[$i]
            $line +=,$tmpline
            $i++
        }
        $line | Out-File $csvFilePth

        #ショートカット作成（Mock_Login_ArgumentSpecial）
        $tmpline=$null
        $tmpline=$sCsvKeys[$nArgSpecialIdx]+","+$sCsvAppName[$nArgSpecialIdx]+","+$sCsvVals[$nArgSpecialIdx]+","+$sCsvOther[$nArgSpecialIdx]
        CreateShortcut $tmpline 7

        #ショートカット作成（SettingMenu）
        CreateShortcut "-WindowStyle Hidden -command SettingMenu,Menu.ps1,,Ctrl+Shift+Alt+A" 1

        #Tipsメッセージを表示
        [System.Windows.Forms.MessageBox]::Show("設定が完了しました。"+"`n"+
                                                "「Ctrl+Shift+Alt+Z」でログインツールを起動してください。",
                                                "Tips",
                                                "OK",
                                                "Information")

    }

} catch [Exception] {
    $exp = $error[0].ToString()
    Read-Host $exp
    exit 1
} finally {

}



#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ 画面起動プチメニュー							Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
# アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# シェルオブジェクト生成
$shell = New-Object -ComObject Shell.Application
# IEリスト取得
$ie_list = @($shell.Windows() | where { $_.Name -match "Internet Explorer" })
# 複数ヒットした場合は一番最後を取得する(添え字-1は最後）
$ie = @($ie_list | where { $_.LocationURL -match $V_IPADDR })[-1]

# フォームを作る
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "画面起動プチメニュー"			#タイトル文字列を指定する
$Form.Size = "300,200"						#フォームの大きさを指定する
#$Form.StartPosition = "CenterScreen"		#フォームを画面の中心に表示する
#$form.StartPosition = "Manual"				#フォームを画面の任意の位置に表示する
#$form.Location = "100,100"					#フォームを画面の任意の位置に表示する
#$form.MaximizeBox = $False					#最大化ボタンを表示させない
#$form.MinimizeBox = $False					#最小化ボタンを表示させない
#$form.FormBorderStyle = "Fixed3D"			#境界線のスタイルを設定する
$form.BackColor = "#202746"					#フォームの背景色を指定する
$form.ForeColor = "#969ca8"					#フォームの前景色(＝文字色)を指定する
#$Form.Opacity = 1							#フォームの透明度を指定する（0～1：0.5や0.9など、0に近い程透明）
#$Form.TopLevel = $true 						#起動時は最前面に表示する(他のアプリケーションに隠れる)
#$Form.FormBorderStyle = "None"				#タイトルバーを表示しない(閉じるボタンや境界線も消える)
#$form.ControlBox = $False					#タイトルバーを表示しない(境界線は残る)
#$form.Text = ""							#タイトルバーを表示しない(境界線は残る)
#$form.ShowInTaskbar = $False				#タスクバーにフォームを表示しない


# 開発統合環境グループを作る
$MyGroupBox = New-Object System.Windows.Forms.GroupBox
$MyGroupBox.Location = "10,10"
$MyGroupBox.size = "265,100"
$MyGroupBox.text = "開発統合環境"

# 開発統合環境グループの中のラジオボタンを作る
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Location = "20,20"
$RadioButton1.size = "200,20"
$RadioButton1.Checked = $True
$RadioButton1.Text = "画面モック"

# OKボタンを作る
$OKButton = new-object System.Windows.Forms.Button
$OKButton.Location = "50,125"
$OKButton.Size = "80,30"
$OKButton.Text = "遷移"
$OKButton.DialogResult = "OK"

# キャンセルボタンを作る
$CancelButton = new-object System.Windows.Forms.Button
$CancelButton.Location = "150,125"
$CancelButton.Size = "80,30"
$CancelButton.Text = "キャンセル"
$CancelButton.DialogResult = "Cancel"

# グループにラジオボタンを入れる
$MyGroupBox.Controls.AddRange(@($Radiobutton1))

# フォームに各アイテムを入れる
$Form.Controls.AddRange(@($MyGroupBox,$OKButton,$CancelButton))

# Enterキー、Escキーと各ボタンの関連付け
$Form.AcceptButton = $OKButton
$Form.CancelButton = $CancelButton


# 無限
do {
	# フォームを表示させ、押されたボタンの結果を受け取る
	$dialogResult = $Form.ShowDialog()

	# ボタン押下による条件分岐
	if ($dialogResult -eq "OK"){
			[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "CommonKyotoTheme_wt88_block_wt45_wtTabsPlaceHolder_wtFunctionListPlaceHolder_wtMenuBtnsList_Dummy02_ctl00_CommonKyotoParts_wt38_block_wtFunctionLink").click()
	}
}while ($dialogResult -eq "OK")

exit

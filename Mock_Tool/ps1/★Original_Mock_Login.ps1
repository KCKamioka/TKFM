#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ��ʃ��b�N�������O�C��																						Ver.1.0
#_/ 
#_/ �y�����z
#_/ Args[0]								Args[1]			Args[2]				Args[3]		Args[4]			Args[5]
#_/ �ڑ��T�[�o�[						�V�X�e��		����				�[��		���t���b�V��	�p�[�\�i���G���A
#_/ --------------------------------------------------------------------------------------------------------------------
#_/ 1�F��Q�ҕ����@���s�s��			9�F��Q�ҕ���	0�F��Q�Ҏ蒠		 0�FPC0010	0�FOff			0�FOff
#_/ 2�F��Q�ҕ����@�J����(114��)					1�F�⑕��			 1�FPC0012	1�FOn			1�FOn
#_/ 3�F��Q�ҕ����@���r����(122��)					2�F��Q�ґ����x��	 2�FPC0014
#_/ 4�F���ہ@���s�s��									3�F�X�����			 3�FPC0016
#_/ 5�F���ہ@�J����									4�F��Q�ҕ����E����	 4�FPC0020
#_/ 6�F���ہ@���r����									5�F���ʏ�Q�Ҏ蓖	 5�FPC0022
#_/ 7�F���@���s�s��									6�F�S�g�}�{����		 6�FPC0024
#_/ 8�F���@�J����									7�F������			 7�FPC0026
#_/ 9�F���@���r����									8�F�Ǘ���			 8�FPC0028
#_/ 																		 9�FPC0030
#_/ 																		10�FPC0032
#_/ 																		11�FPC0034
#_/ 																		12�FPC0036
#_/ 																		13�FPC0038
#_/ 																		14�FPC0040
#_/ 																		15�FPC0099
#_/ 
#_/ 																										SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing
#�n�b�V��
$settings = @{}
#�C���|�[�g�������ڂ��n�b�V���ϐ��Ƀ}�b�s���O���Ȃ���i�[
Import-Csv -Path ..\Config\setting.ini -Header Key,Value -Delimiter "=" | %{$settings.Add($_.Key.Trim(),$_.Value.Trim())}

#�t�q�k
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
  default {Write-Output '�ϐ�$Args[0]��1�`9�ł͂���܂���B'}
}

if($Args[5] -eq "1") {
	#�p�[�\�i���G���A�̐ݒ�
	$url=$url+$settings.PERSONAL_AREA
}

#��ʓ��͒l
$strUser		= $settings.MOCK_USER
$strPass		= $settings.MOCK_PASS
$strSystem		= $Args[1]
$strBusiness	= $Args[2]
$strTerminalId	= $Args[3]

#IE�N��
$ie = New-Object -ComObject InternetExplorer.Application

# IE��ʒ�����������
$ie.top			= $settings.WINDOW_TOP
$ie.left		= $settings.WINDOW_LEFT
$ie.width		= 600
$ie.height		= 400
$ie.StatusBar	= $false
$ie.ToolBar		= $false
$ie.MenuBar		= $false
$ie.AddressBar	= $false
$ie.Resizable	= $false

# BLANK�y�[�W���J���āC�E�B���h�E���T�C�Y���擾����
$ie.Navigate("about:blank")
$ie.Document.IHTMLDocument2_write(@"
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja" lang="ja">
<meta http-equiv="Content-Type" content="text/html;charset=Shift_JIS">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<title>JavaScript �e�X�g</title>
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

# �y�[�W�����[�h���I���܂ő҂�
while ($ie.Busy) {sleep -milliseconds 500}

# �E�B���h�E���T�C�Y��1280*1002�ɂ���
$test = $ie.Document.IHTMLDocument2_body.innerText
$size = $test.Split(":")
$ie.width=600+(1280-$size[0])
$ie.height=400+(1002-$size[1])

# �ړI��URL���J��
$ie.Navigate($url)

# �y�[�W�����[�h���I���܂ő҂�
while ($ie.Busy) {sleep -milliseconds 500}

# IE����������
$ie.Visible = $true

# IE���A�N�e�B�u�ɂ���
add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
$b = (Get-Process | Where-Object {$_.MainWindowTitle -like "*Internet Explorer*"}).id
start-sleep -Milliseconds 500
[Microsoft.VisualBasic.Interaction]::AppActivate($b)

# �ؖ����ł̑��s
#$element = [System.__ComObject].InvokeMember(
#			"getElementById"										# �w����@��ID���g�p
#			,[System.Reflection.BindingFlags]::InvokeMethod			# ���܂��Ȃ�
#			, $null													# �s�v�p�����[�^
#			, $ie.Document											# �y�[�W���I�u�W�F�N�g
#			, "overridelink"										# ID��
#			)
#if($element -ne [System.DBNull]::Value) {
#	$element.click()
#}
#
# �y�[�W�����[�h���I���܂ő҂�
#while ($ie.Busy) {sleep -milliseconds 500}

# ���O�C�����ݒ�
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtUsernameInput").value=$strUser
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtPasswordInput").value=$strPass
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
start-sleep -milliseconds 1000
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtBusinessIdComboBox").SelectedIndex=$strBusiness
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtTerminalIdComboBox").SelectedIndex=$strTerminalId
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtLoginButton").click()

# �y�[�W�����[�h���I���܂ő҂�
while ($ie.Busy) {sleep -milliseconds 1000}
start-sleep -Milliseconds 1000

if($Args[4] -eq "1") {
	start-sleep -milliseconds 3000
	#���t���b�V���N��
	start-process powershell.exe -windowstyle hidden .\Mock_reflesh.ps1
	#���j�^�����O�N��
	start-process powershell.exe -windowstyle hidden .\Mock_monitoring.ps1
}

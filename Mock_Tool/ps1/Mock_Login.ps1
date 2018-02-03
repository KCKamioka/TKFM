#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ��ʃ��b�N�������O�C��																						Ver.1.0
#_/ 
#_/ �y�����z
#_/ Args[0]							Args[1]				Args[2]				Args[3]		Args[4]			Args[5]
#_/ �ڑ��T�[�o�[					�V�X�e��			����				�[��		���t���b�V��	�p�[�\�i���G���A
#_/ --------------------------------------------------------------------------------------------------------------------
#_/ 1�F�Z��@�@�@�J����			10�F�������		0�F�V�l���			 0�FPC0010	0�FOff			0�FOff
#_/ 2�F�Z��@�@�@�����J����		11�F�h�V��ԏ�		1�F��Q�҈��		 1�FPC0012	1�FOn			1�FOn
#_/ 3�F������Á@�J����			12�F�Z��E�ŏƉ�	2�F�ЂƂ�e���		 2�FPC0014
#_/ 4�F������Á@�����J����		13�F����			3�F�q�ǂ����		 3�FPC0016
#_/ 													4�F������			 4�FPC0020
#_/ 													5�F�Ǘ���			 5�FPC0022
#_/ 																		 6�FPC0024
#_/ 																		 7�FPC0026
#_/ 																		 8�FPC0028
#_/																			 9�FPC0030
#_/ 																		10�FPC0032
#_/ 																		11�FPC0034
#_/ 																		12�FPC0036
#_/ 																		13�FPC0038
#_/ 																		14�FPC0040
#_/ 																		15�FPC0088
#_/ 																		16�FPC0099
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
  1 {$url="http://"+$settings.JUKI_DEVELOP_SERVER+$settings.URL_VALUE1}			#192.168.60.24
  2 {$url="https://"+$settings.JUKI_REVIEW_SERVER+$settings.URL_VALUE1}			#juki-dc.open.kyoto.jp
  3 {$url="http://"+$settings.FUKUSHI_DEVELOP_SERVER+$settings.URL_VALUE1}		#192.168.60.21
  4 {$url="https://"+$settings.FUKUSHI_REVIEW_SERVER+$settings.URL_VALUE1}		#fukui-dc.open.kyoto.jp
  default {Write-Output '�ϐ�$Args[0]��1�`4�ł͂���܂���B'}
}

if($Args[5] -eq "1") {
	#�p�[�\�i���G���A�̐ݒ�
	$url=$url+$settings.PERSONAL_AREA
}

#��ʓ��͒l

# �u������Ái�����J���j�v�̏ꍇ���[�U�[�E�p�X�͌Œ�
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

#IE�N��
$ie = New-Object -ComObject InternetExplorer.Application

# IE��ʒ�����������
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
$ie.width=300+(1280-$size[0])
$ie.height=200+(1002-$size[1])

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

# ���O�C�����ݒ�
# �u�Z��v�̏ꍇ���[�U�[�E�p�X�͐ݒ肵�Ȃ�
if($Args[1] -ge "3") {
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtUsernameInput").value=$strUser
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtPasswordInput").value=$strPass
}

# # �u������Ái�J�����j�v�̏ꍇ�{�^���w��
# if($Args[0] -eq "3") {
#     #�V�X�e��
#     switch($Args[1]){
#         10 {
#             #����
#             switch($Args[2]){
#                 0 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt60").click()}		#�V�l���
#                 1 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt20").click()}		#��Q�҈��
#                 2 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt43").click()}		#�ЂƂ�e���
#                 3 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt56").click()}		#�q�ǂ����
#                 default {Write-Output '�ϐ�$Args[0]��1�`4�ł͂���܂���B'}
#             }
#         }		#�������
#         11 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt118").click()}		#�h�V��ԏ�
#         12 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wt46").click()}		#�Z��E�ŏƉ�
#         13 {[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
#             sleep -milliseconds 500
#             [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
#             }		#����
#         default {Write-Output '�ϐ�$Args[0]��1�`4�ł͂���܂���B'}
#     }
#     sleep -milliseconds 1000
# }else{
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").SelectedIndex=$strSystem
    sleep -milliseconds 500
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtSystemIdComboBox").FireEvent("onchange")
    sleep -milliseconds 500
    [System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtBusinessIdComboBox").SelectedIndex=$strBusiness
# }

# �[��
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtTerminalIdComboBox").SelectedIndex=$strTerminalId
# ���O�C��
[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "wtLoginButton").click()

# �y�[�W�����[�h���I���܂ő҂�
while ($ie.Busy) {sleep -milliseconds 3000}

if($Args[4] -eq "1") {
	start-sleep -milliseconds 3000
	#���t���b�V���N��
	start-process powershell.exe -windowstyle hidden .\Mock_reflesh.ps1
	#���j�^�����O�N��
	start-process powershell.exe -windowstyle hidden .\Mock_monitoring.ps1
}

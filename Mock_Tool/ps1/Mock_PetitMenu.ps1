#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ��ʋN���v�`���j���[							Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
# �A�Z���u���̃��[�h
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �V�F���I�u�W�F�N�g����
$shell = New-Object -ComObject Shell.Application
# IE���X�g�擾
$ie_list = @($shell.Windows() | where { $_.Name -match "Internet Explorer" })
# �����q�b�g�����ꍇ�͈�ԍŌ���擾����(�Y����-1�͍Ō�j
$ie = @($ie_list | where { $_.LocationURL -match $V_IPADDR })[-1]

# �t�H�[�������
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "��ʋN���v�`���j���["			#�^�C�g����������w�肷��
$Form.Size = "300,200"						#�t�H�[���̑傫�����w�肷��
#$Form.StartPosition = "CenterScreen"		#�t�H�[������ʂ̒��S�ɕ\������
#$form.StartPosition = "Manual"				#�t�H�[������ʂ̔C�ӂ̈ʒu�ɕ\������
#$form.Location = "100,100"					#�t�H�[������ʂ̔C�ӂ̈ʒu�ɕ\������
#$form.MaximizeBox = $False					#�ő剻�{�^����\�������Ȃ�
#$form.MinimizeBox = $False					#�ŏ����{�^����\�������Ȃ�
#$form.FormBorderStyle = "Fixed3D"			#���E���̃X�^�C����ݒ肷��
$form.BackColor = "#202746"					#�t�H�[���̔w�i�F���w�肷��
$form.ForeColor = "#969ca8"					#�t�H�[���̑O�i�F(�������F)���w�肷��
#$Form.Opacity = 1							#�t�H�[���̓����x���w�肷��i0�`1�F0.5��0.9�ȂǁA0�ɋ߂��������j
#$Form.TopLevel = $true 						#�N�����͍őO�ʂɕ\������(���̃A�v���P�[�V�����ɉB���)
#$Form.FormBorderStyle = "None"				#�^�C�g���o�[��\�����Ȃ�(����{�^���⋫�E����������)
#$form.ControlBox = $False					#�^�C�g���o�[��\�����Ȃ�(���E���͎c��)
#$form.Text = ""							#�^�C�g���o�[��\�����Ȃ�(���E���͎c��)
#$form.ShowInTaskbar = $False				#�^�X�N�o�[�Ƀt�H�[����\�����Ȃ�


# �J���������O���[�v�����
$MyGroupBox = New-Object System.Windows.Forms.GroupBox
$MyGroupBox.Location = "10,10"
$MyGroupBox.size = "265,100"
$MyGroupBox.text = "�J��������"

# �J���������O���[�v�̒��̃��W�I�{�^�������
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Location = "20,20"
$RadioButton1.size = "200,20"
$RadioButton1.Checked = $True
$RadioButton1.Text = "��ʃ��b�N"

# OK�{�^�������
$OKButton = new-object System.Windows.Forms.Button
$OKButton.Location = "50,125"
$OKButton.Size = "80,30"
$OKButton.Text = "�J��"
$OKButton.DialogResult = "OK"

# �L�����Z���{�^�������
$CancelButton = new-object System.Windows.Forms.Button
$CancelButton.Location = "150,125"
$CancelButton.Size = "80,30"
$CancelButton.Text = "�L�����Z��"
$CancelButton.DialogResult = "Cancel"

# �O���[�v�Ƀ��W�I�{�^��������
$MyGroupBox.Controls.AddRange(@($Radiobutton1))

# �t�H�[���Ɋe�A�C�e��������
$Form.Controls.AddRange(@($MyGroupBox,$OKButton,$CancelButton))

# Enter�L�[�AEsc�L�[�Ɗe�{�^���̊֘A�t��
$Form.AcceptButton = $OKButton
$Form.CancelButton = $CancelButton


# ����
do {
	# �t�H�[����\�������A�����ꂽ�{�^���̌��ʂ��󂯎��
	$dialogResult = $Form.ShowDialog()

	# �{�^�������ɂ���������
	if ($dialogResult -eq "OK"){
			[System.__ComObject].InvokeMember("getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ie.Document, "CommonKyotoTheme_wt88_block_wt45_wtTabsPlaceHolder_wtFunctionListPlaceHolder_wtMenuBtnsList_Dummy02_ctl00_CommonKyotoParts_wt38_block_wtFunctionLink").click()
	}
}while ($dialogResult -eq "OK")

exit

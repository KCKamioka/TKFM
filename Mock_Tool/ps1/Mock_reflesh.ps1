#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ��ӂ������									Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing

#���s�v���O����
$execPath = '.\Mock_monitoring.lnk'
#���s�v���O�����N��
start-process $execPath

# �G���[�����ŏI��
$ErrorActionPreference = "stop"

# �V�F���I�u�W�F�N�g����
$shell = New-Object -ComObject Shell.Application
# IE���X�g�擾
$ie_list = @($shell.Windows() | where { $_.Name -match "Internet Explorer" })
# �����q�b�g�����ꍇ�͈�ԍŌ���擾����(�Y����-1�͍Ō�j
$ie = @($ie_list | where { $_.LocationURL -match $V_IPADDR })[-1]

# IE��������Ȃ��Ȃ�ƃG���[�ɂȂ�̂�Shell�͏I������
do {
	#�uStart-Sleep -s 300�v��300�b(5��)�X���[�v
	Start-Sleep -s 300
	# IE�����t���b�V��
	$ie.Refresh()
}while ($true)

exit

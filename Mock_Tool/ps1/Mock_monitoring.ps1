#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ ���ɂ����									Ver.1.0
#_/ 
#_/												SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Add-Type -AssemblyName System.Drawing

#�Ď��v���Z�X��
$procName = "iexplore"

#�I���v���Z�X��
$EndProc = "powershell"

#�Ď�&�����N��
try{
	#�v���Z�X���擾
	$p_lis = Get-Process $procName -ErrorAction SilentlyContinue
	# �����q�b�g�����ꍇ�͈�ԍŌ���擾����(�Y����-1�͍Ō�j
	$p = @($p_lis | where { $_.LocationURL -match $V_IPADDR })[-1]

	#�J�n�����v���Z�X���I������܂őҋ@����悤�Ɏw��
	$p.WaitForExit()
}
catch{
	Write-Host "No Process"
}
finally{
	#�v���Z�X�I��
	Stop-process -name $EndProc
}

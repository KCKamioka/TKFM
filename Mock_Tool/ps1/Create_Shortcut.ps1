#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
#_/ �V���[�g�J�b�g�쐬								Ver.1.0
#_/ 
#_/ 											SystemTKFM
#_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
try{
	# �X�N���v�g�����݂���f�B���N�g�����擾
	$base = split-Path -Parent $MyInvocation.MyCommand.Path;

	#�V�F���I�u�W�F�N�g�쐬
	$ws = New-Object -com WScript.Shell;

	# �����t�H���_������ꍇ�A�����̃t�H���_�͉����ς��Ȃ�
	New-Item "..\MockShortcut" -type directory -Force 

	# �uMock_Tool.csv�v��ǂݍ��݁A�V���[�g�J�b�g���쐬����
	foreach ($line in (get-content "$base`\..\Config\Mock_Tool.csv")){
		$s = $line.split(",");
		#�V���[�g�J�b�g�I�u�W�F�N�g
		$lnk = $ws.CreateShortcut($base + "\..\MockShortcut\" + $s[0] + ".lnk");
		#�����N��(�ΏۂƂȂ�t�@�C��)
		$lnk.TargetPath = "powershell";
		#�R�}���h���C������
		$lnk.Arguments = $base+ "\" + $s[1] + " " + $s[2];
		#��ƃt�H���_
		$lnk.WorkingDirectory = $base;
		#�A�C�R��(�p�X+�C���f�b�N�X)
		$lnk.IconLocation = $s[5] + "," + $s[6];
		#�V���[�g�J�b�g�L�[( "Alt+", "Ctrl+", "Shift+" �ȂǂƋ���)
		$lnk.Hotkey = $s[3];
		#���s���̑傫��(1:�ʏ�E�C���h�E�A3:�ő剻�A7:�ŏ���)
		$lnk.WindowStyle = "7";
		#�R�����g(�V���[�g�J�b�g�̐���)
		$lnk.Description = $s[4];
		#�v���p�e�B�̕ۑ�
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

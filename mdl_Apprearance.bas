Attribute VB_Name = "mdl_Apprearance"
'# �A�v���P�[�V�����̕\��

''## ���{���̕\��
''- �\��
'   Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
''- ��\��
'   Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

''## �X�e�[�^�X�o�[�̕\��
''- �\��
'   Application.DisplayStatusBar = True
''- ��\��
'   Application.DisplayStatusBar = False

''## �����o�[�̕\��
''- �\��
'    Application.DisplayFormulaBar = True
''- ��\��
'    Application.DisplayFormulaBar = False

''## �^�u�̕\��
''- �\��
'    ActiveWindow.DisplayHeadings = True
''- ��\��
'    ActiveWindow.DisplayHeadings = False

''## ��ʂ��őO�ʂɕ\������
'    AppActivate Application.Caption

''## ��ʂ̃T�C�Y
''- �ő剻����
'    Application.WindowState = xlMaximized
'- �ŏ�������
'    Application.WindowState = xlMinimized
'- �W���̑傫���ɂ���
'    Application.WindowState = xlNormal

''## �A�v���P�[�V�����G���[�̔��o
''- ���o����
'    Application.DisplayAlerts = True
''- ���o���Ȃ�
'    Application.DisplayAlerts = False

''## ��ʍX�V�ݒ�
''- �X�V����
'    Applicaiton.ScreenUpdating = True
''- �X�V���Ȃ�
'    Application.ScreenUpdating = False

''## �؂���@�\�̐ݒ�
''- �L����
'    Application.CutCopyMode = 2
''- ������
'    Application.CutCopyMode = 0

''## �C�x���g�̗}�~
''- �}�~����
'    Application.EnableEvents = False
''- �}�~
'    Application.EnableEvents = True

''## �V�[�g�̕\��/��\��
''- �\��
'    Sheets("Sheet1").Visible = xlVisible
''- ��\���i�V�[�g�^�u��ōĕ\���ݒ�\�j
'    Sheets("Sheet1").Visible = xlHidden
''- ���S�ɔ�\���iVBE�̃v���p�e�B�ɂĕ\�����[�h�ɕύX�\�j
'    Sheets("Sheet1").Visible = xlVeryHidden

''## �V�[�g�̕ی�
''- �ی�
'    Sheets("Sheet1").Protect
''- �ی����
'    Sheets("Sheet1").Unprotect

''## �V���[�g�J�b�g�L�[�̐ݒ�
''- ������
'    Application.OnKey "%{F11}", ""
'    Application.OnKey "^x", ""
'    Application.OnKey "^s", ""
''- �L����
'    Application.OnKey "%{F11}"
'    Application.OnKey "^x"
'    Application.OnKey "^s"
''- �V���[�g�J�b�g�L�[�̖����ύX
'    Application.OnKey "^x", "^s"



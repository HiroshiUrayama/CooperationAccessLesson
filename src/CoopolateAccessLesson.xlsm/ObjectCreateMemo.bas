Attribute VB_Name = "ObjectCreateMemo"
'#################################################
'CreateObject�֐��ɓ����N���X��
'#################################################
Option Explicit

'CreateObject(class,[servername])
'class�ɂ́A�uappname.objecttype�v�Ƃ����\���ō쐬����I�u�W�F�N�g�̃A�v���P�[�V�������ƃN���X���L�ڂ���
'servername�͏ȗ�����邩�󕶎���""�̏ꍇ�ɂ̓��[�J���R���s���[�^���g����B
'---------------------------------------------------
'CreateObject�֐��Ɏg����ProgID�́A
'���W�X�g�� HKEY_LOCAL_MACHINE\SOFTWARE\Classes�ɕۑ�����Ă���B
'---------------------------------------------------

Sub createObjectStates()
    '�Q�Ɛݒ�Microsoft.Access XX.X Object Library
    Dim a As Object
    Set a = CreateObject("Access.Application")
    
    'DAO�̃f�[�^�x�[�X�I�u�W�F�N�g
    'accdb�t�@�C���̏ꍇ��Microsoft Office XX.X Access Database Engine Object Library
    'mdb�t�@�C���̏ꍇ��Microsoft DAO 3.6 Object Library
    '64bit�o�[�W������Access�ł͎Q�ƕs�\
    Dim b As Object
    Set b = CreateObject("DAO.Database")
    
    'ADO�̐ڑ�������(Microsoft ActiveX Data Objects X.X Library)
    Dim c As Object
    Set c = CreateObject("ADODB.Connection")
    
    'ADO�̃��R�[�h�Z�b�g(Microsoft ActiveX DataObjects X.X Library)
    Dim d As Object
    Set d = CreateObject("ADODB.Recordset")
    
    'ADO�̃X�g���[���I�u�W�F�N�g(Microsoft ActiveX Data Objects X.X Library)
    Dim e As Object
    Set e = CreateObject("ADODB.Stream")
    
    'ADOX�̃J�^���O�I�u�W�F�N�g�B
    '�f�[�^�x�[�X�ɐڑ����鎞�Ɏw��(Microsoft ADOExt. X.X for DDLand Security)
    Dim f As Object
    Set f = CreateObject("ADOX.Catalog")
    
    'Excel�̃A�v���P�[�V����(Microsoft Excel XX.X Object Library)
    Dim g As Object
    Set g = CreateObject("Excel.Application")
    
    'Excel�̃��[�N�V�[�g(�hMicrosoft Excel XX.X Object Library)
    Dim h As Object
    Set h = CreateObject("Excel.Worksheet")
    
    'Outlook�̃A�v���P�[�V����(Microsoft Excel XX.X Object Library)
    Dim i As Object
    Set i = CreateObject("Outlook.Application")
    
    'PowerPoint�̃A�v���P�[�V����(Microsoft PowerPoint XX.X Object Library)
    Dim j As Object
    Set j = CreateObject("PowerPoint.Application")
    
    'Word�̃A�v���P�[�V����(Microsoft Word XX.X Object Library)
    Dim k As Object
    Set k = CreateObject("Word.Application")
    
    'InternetExplorer(Microsoft Internet Controls)
    Dim l As Object
    Set l = CreateObject("ADODB.Stream")
    
    '�t�@�C���V�X�e���I�u�W�F�N�g(Microsoft Scripting Runtime)
    Dim m As Object
    Set m = CreateObject("Scripting.FileSystemObject")
    
    '�f�B�N�V���i���I�u�W�F�N�g("Microsoft Scripting Runtime")
    Dim n As Object
    Set n = CreateObject("Scripting.Dictionary")
    
    'WindowsScriptHost�̃V�F���I�u�W�F�N�g(Windows Script Host Object Model)
    Dim o As Object
    Set o = CreateObject("WScript.Shell")
    
    'Windows�̃V�F���I�u�W�F�N�g(Microsoft Shell Controls And Automation)
    Dim p As Object
    Set p = CreateObject("Shell.Application")
    
    '���K�\���Ŏg�p����I�u�W�F�N�g(Microsoft VBScript Regular Expressions 5.5)
    Dim q As Object
    Set q = CreateObject("VBScript.RegExp")
    
End Sub

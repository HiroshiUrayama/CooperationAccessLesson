Attribute VB_Name = "OtherApplication"
'#################################################
'���̃A�v���P�[�V�����Ƃ̘A�g
'#################################################
Option Explicit

'==================================================
'Outlook����������_�������Ă����ׂ��̓I�u�W�F�N�g���f��
'==================================================
'outlook�̓��[�������łȂ��A�X�P�W���[���Ǘ����ł��邩��

'Outlook�̃I�u�W�F�N�g���f��
'---------------------------------------------------
'�yApplication�I�u�W�F�N�g�z
'Outlook���̂��̂�\���B�������AExcel�܂ߑ���Office�A�v���P�[�V�����ƈقȂ�A������Aplicaton�I�u�W�F�N�g���쐬�ł��Ȃ��B
'���̂��߁A����Outlook���N�����Ă����ԂŁA�V����Application�I�u�W�F�N�g���쐬���Ă��A
'�ʂ�Outlook���N������񂶂�Ȃ��āA�N������Outlook���Q�Ƃ���B

'�Ȃ��AExcel�ł́AExcel��Workbook�I�u�W�F�N�g��V�K�쐬����ƁA�����I�ɃA�v���P�[�V�������N������(�����I��Application�I�u�W�F�N�g���Q�Ƃ��Ȃ��Ă�)�B
'�������AOutlook�͍ŏ���Application�I�u�W�F�N�g�𐶐�����K�v������B

'�yNamespace�I�u�W�F�N�g�z
'�A�h���X���⃁�b�Z�[�W�Ȃǂ̃f�[�^�ɃA�N�Z�X���邽�߂̃C���^�[�t�F�[�X�B
'Outlook��VBA�ň����ꍇ�̓�����B
'NameSpace�I�u�W�F�N�g�́AApplication�I�u�W�F�N�g��GetNameSpace���\�b�h�ň�����"MAPI"���w�肵�Ď擾����B
'�Ȃ��AGetNameSpace���\�b�h�̈����Ɏw��ł���l��"MAPI"�̂݁B

'��MAPI
'Windows��Microsoft�҂̃A�v���P�[�V��������d�q���[�����������߂̕W���d�l�̂��Ƃ������B

'�yExplorer�I�u�W�F�N�g�z
'�t�H���_��\�����Ă���E�C���h�E�ɑΉ�����I�u�W�F�N�g�B
'�t�H���_�̃r���[��ǂ̃A�C�e����I�����Ă��邩�ƌ��������Ȃǂ��擾�E�ݒ肷��B
'Application��Explorers�R���N�V�����܂���ActiveExplorer�v���p�e�B�ŁAExplorer�I�u�W�F�N�g�擾�ł���B

'�yInspector�I�u�W�F�N�g�z
'�A�C�e���̕\�����s���Ă���E�C���h�E�ɑΉ�����I�u�W�F�N�g�B
'�\�����̃A�C�e����A�C�e���E�C���h�E�̃R�}���h�{�^�����̏����擾�E�ݒ肷��B
'Application�I�u�W�F�N�g��Inspectors�R���N�V�����܂���ActiveInspector�v���p�e�B�ɂ���āAInspector�I�u�W�F�N�g���擾���邱�Ƃ��ł���B

'�yFolder�I�u�W�F�N�g�z
'�t�H���_�ɑΉ�����I�u�W�F�N�g�B���̃I�u�W�F�N�g��ʂ��āA�t�H���_���̃��b�Z�[�W��T�u�t�H���_�A�r���[�Ȃǂ̏��ɃA�N�Z�X����B
'��M�g���C�⑗�M�ς݃A�C�e���ȂǁA
'Outlook�N�����ɍ쐬�����K��̃t�H���_�ɂ��ẮA
'NameSpace�I�u�W�F�N�g��GetDefaultFolder���\�b�h���g�p���Ď擾�\�B

'�yItem�I�u�W�F�N�g�z
'�t�H���_�Ɋi�[����Ă���A�C�e���ɂ́A���ꂼ��̃A�C�e���̎�ʂɉ������I�u�W�F�N�g������B
'�Ⴆ�΁A���b�Z�[�W�̃A�C�e����MailItem�I�u�W�F�N�g�A�\���AppointmentItem�I�u�W�F�N�g�A�d����TaskItem�I�u�W�F�N�g�ƂȂ�B
'Item�I�u�W�F�N�g�́AFolder�I�u�W�F�N�g��Items�R���N�V��������擾�ł���B
'�܂��AApplication�I�u�W�F�N�g��CreateItem���\�b�h�ŐV�K�A�C�e�����쐬���邱�Ƃ��\�B

'Outlook�ɒ��ڏ�����VBA�̏���
'---------------------------------------------------
'Excel�݂����ɔz�z���Ďg�����Ƃ��ł��Ȃ��B
    '�����W���[����z�z���Ďg�p����l��Outlook��Code���R�s�[����
    'Module���C���|�[�g����


'==================================================
'�V�K���[�����쐬����
'==================================================
Private Sub CreateNewMail()
    Dim olApp As Outlook.Application
    Dim MailItem As Outlook.MailItem
    
    'Outlook���N������
    Set olApp = New Outlook.Application
        
    '���[�����쐬����
    Set MailItem = olApp.CreateItem(olMailItem)
    
    With MailItem
        '���M����w�肷��
        .Recipients.Add("hiroshiurayama0308@gmail.com").Type = 1
        .Subject = "VBA���瑗�M"
        
        '�{���̐ݒ�
        .Body = "������" & vbCrLf & _
                    "����ɂ��́B�����ł��B" & vbCrLf & _
                    "���K�Ŏ����ɑ��M���Ă��܂��B" & vbCrLf & _
                    "�����͐����̉J�ł����A�̒��͂������ł��傤��?" & vbCrLf & _
                    "�������{���l�[�[�΂���ő̒��������Ȃ肻���Ȃ̂ŁA" & vbCrLf & _
                    "�����͂Ȃ߂��̂����X�`�Ƃ��A�����̏��Ȃ����̂�H�ׂĉ񕜂��܂��傤�ˁB" & vbCrLf & _
                    "�ȏ�ł��B"
                    
        '�Y�t�t�@�C�����w�肷��
        .Attachments.Add "C:\Users\USER\Desktop\LineUp�p_�ʐ^�A����\�ʐ^\IMG_20210213_152036.jpg"
        '.Send
        .Display
        
        .Send
    End With
End Sub

'==================================================
'���[���̏����擾����T���v��
'==================================================
Private Sub GetMailItem()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.Folder
    Dim i As Long
    
    'outlook���N������
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    '�Ώۃt�H���_��I������
    Set olFolder = olNamespace.PickFolder
    
    '�S�ẴA�C�e���ɑ΂��ď������s��
    For i = 1 To olFolder.Items.Count
        If olFolder.Items(i).Class = olMail Then
            Cells(i, 1).Value = olFolder.Items(i).SenderName
            Cells(i, 2).Value = olFolder.Items(i).Subject
            Cells(i, 3).Value = olFolder.Items(i).ReceivedTime
            Cells(i, 4).Value = olFolder.Items(i).Body
            'Debug.Print "---------------------------------------"
        End If
    Next

End Sub






































































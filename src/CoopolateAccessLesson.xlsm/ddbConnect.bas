Attribute VB_Name = "ddbConnect"
'#################################################
'�f�[�^�x�[�X�Ƃ̘A�g
'#################################################
Option Explicit

'==================================================
'�ڋq�f�[�^.accdb��T�ڋq���X�g�e�[�u���̃t�B�[���h�����o�͂���
'==================================================
Private Sub DAOSample()
    
    '���[�N�X�y�[�X
    Dim ws As DAO.Workspace
    
    'DBEngine���ŏ�ʃI�u�W�F�N�g��(�f�[�^�x�[�X����Ƃ���̈�)���[�N�X�y�[�X���`����
     Set ws = DAO.DBEngine.Workspaces(0)
     
     '�f�[�^�x�[�X�ւ̎Q��
    Dim db As DAO.Database
    
    '�ڋq�f�[�^.accdb�ɐڑ�����
    Set db = ws.OpenDatabase(ThisWorkbook.Path & "\�ڋq�f�[�^.accdb")
     
     '���R�[�h�Z�b�g
    Dim rs As DAO.Recordset
    
    'T�ڋq���X�g�e�[�u���̃��R�[�h�Z�b�g���J��
     Set rs = db.OpenRecordset("T�ڋq���X�g", dbOpenDynaset)
     
    Dim i As Long
    For i = 0 To rs.Fields.Count - 1    'Fields�R���N�V�����𗘗p���āAField�I�u�W�F�N�g���擾
        '�t�B�[���h�����o�͂���
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

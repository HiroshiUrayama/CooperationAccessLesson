Attribute VB_Name = "DataControl"
'#################################################
'SQL
'#################################################
Option Explicit

'==================================================
'�N�G�������s����R�[�h
'MicrosoftDAO3.6�I�u�W�F�N�g���L���ɂȂ��Ă����Microsoft 16.0 Data(DAO)�Ɨ����ł��Ȃ�(�ǂ�����Q�Ƃ���̂�"DAO�Ȃ̂�")
'==================================================
Private Sub DAOQuerySample()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long
    
    Set ws = DAO.DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(ThisWorkbook.Path & "\�ڋq�f�[�^.accdb")
    Set rs = db.OpenRecordset("Q�����s")
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

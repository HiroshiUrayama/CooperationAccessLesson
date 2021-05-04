Attribute VB_Name = "DataControl"
'#################################################
'SQL
'#################################################
Option Explicit

'==================================================
'クエリを実行するコード
'MicrosoftDAO3.6オブジェクトが有効になっているとMicrosoft 16.0 Data(DAO)と両立できない(どちらも参照するのは"DAOなので")
'==================================================
Private Sub DAOQuerySample()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long
    
    Set ws = DAO.DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(ThisWorkbook.Path & "\顧客データ.accdb")
    Set rs = db.OpenRecordset("Q東京都")
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

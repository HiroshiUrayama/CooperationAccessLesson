Attribute VB_Name = "ddbConnect"
'#################################################
'データベースとの連携
'#################################################
Option Explicit

'==================================================
'顧客データ.accdbのT顧客リストテーブルのフィールド名を出力する
'==================================================
Private Sub DAOSample()
    
    'ワークスペース
    Dim ws As DAO.Workspace
    
    'DBEngineが最上位オブジェクト→(データベースを作業する領域)ワークスペースを定義する
     Set ws = DAO.DBEngine.Workspaces(0)
     
     'データベースへの参照
    Dim db As DAO.Database
    
    '顧客データ.accdbに接続する
    Set db = ws.OpenDatabase(ThisWorkbook.Path & "\顧客データ.accdb")
     
     'レコードセット
    Dim rs As DAO.Recordset
    
    'T顧客リストテーブルのレコードセットを開く
     Set rs = db.OpenRecordset("T顧客リスト", dbOpenDynaset)
     
    Dim i As Long
    For i = 0 To rs.Fields.Count - 1    'Fieldsコレクションを利用して、Fieldオブジェクトを取得
        'フィールド名を出力する
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

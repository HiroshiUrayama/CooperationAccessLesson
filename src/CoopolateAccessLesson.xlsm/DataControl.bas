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
    
    'クエリが何を表していて、どんな処理を行うのかがわからない。
    'Accessを開きに行って、クエリビューを見ないと…
    Set rs = db.OpenRecordset("Q東京都")
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

'==================================================
'SQLを使用した例
'==================================================
Private Sub SQLSample1()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim vSQL As String
    Dim i As Long
    
    'ワークスペースを定義
    Set ws = DAO.DBEngine.Workspaces(0)
    
    '「顧客データ.accdbに接続する」
    Set db = ws.OpenDatabase("C:\Users\USER\Desktop\Git管理下\CooperationAccessLesson" & "\顧客データ.accdb")
    
    '「Q東京都」テーブルのレコードセットを開く
    '抽出するSQL分を代入する
    'SQLだとエディタ見たら何をやろうとしているのかが一発で分かる
    vSQL = "SELECT * FROM T顧客リスト WHERE 都道府県 = '東京都'"
    
    '指定したsql文を使ってレコードセットを開く
    Set rs = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    'レコードセットの全てのフィールドに対して処理を行う
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

'==================================================
'Accessからデータを取得してワークシートに貼付するサンプル
'==================================================
Private Sub SQLSample2()
    
#Const ARI = True

#If ARI Then
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#Else
    Dim cn As Object
    Dim rs As Object
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
#End If

    'コネクションオープン
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & "C:\Users\USER\Desktop\Git管理下\CooperationAccessLesson" & "\顧客データ.accdb"

    '35歳以上のデータを取得するSQL文を設定する
    With rs
        .ActiveConnection = cn
        .Source = "SELECT * FROM T顧客リスト WHERE 年齢 > 35"
        .Open
    End With
    
    'フィールド名を入力する
    Dim i As Long
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1).Value = rs.Fields(i).Name
    Next
    
    '取得したデータをセルA2以降に貼付する
    'これ、便利やな…
    Range("A2").CopyFromRecordset rs
        
    'コネクション閉じる
    cn.Close
    
    '列幅自動調整
    Columns("A:K").AutoFit
    
End Sub

'==================================================
'Accessからデータを取得してワークシートに貼付するサンプル(Update編)
'==================================================
Private Sub UpdateSample()
#Const ado = True

#If ado Then
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
#Else
    Dim cn As Object
    Set cn = CreateObject("ADODB.Conneciton")
#End If
    
    Dim vSQL As String
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & "C:\Users\USER\Desktop\Git管理下\CooperationAccessLesson" & "\顧客データ.accdb"
    
    vSQL = "UPDATE T顧客リスト SET 年齢 = 48 WHERE 顧客名 = '田中 洋行'"

    '設定したSQL文を実行する
    cn.Execute vSQL
    
    cn.Close
    
End Sub

'#If ARI Then
'#Else
'#End If































































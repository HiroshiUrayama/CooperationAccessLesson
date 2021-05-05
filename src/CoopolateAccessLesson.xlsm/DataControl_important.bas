Attribute VB_Name = "DataControl_important"
'#################################################
'高度な処理(データベース編)
'#################################################
Option Explicit

'重要
'’---------------------------------------------------
'データの整合性を保つこと
'ミスでデータが間違った値に更新されないようにすること

'OpenRecordsetメソッド(DAO)の引数LockEditに指定するLockTypeEnum列挙型の値
'==================================================
'dbOptimistic(3)
'レコードIDにもど突く共有的同時ロック。カーソルは古いレコードと新しいレコードのレコードIDを比較し、
'そのレコードへのアクセスが最後に行われてから変更が加えられたかどうか判断する

'dbPessimistic(2)
'排他的同時ロック。
'カーソルは、レコードが更新可能であることを保証するために必要な最低限のロックを使用する。

'Openメソッド(ADO)の引数LockTypeに指定するLockTypeEnum列挙型の値
'==================================================
'adLcokReadOnly(1)
'既定値。読み取り専用。データの変更不可。

'adLockPessimistic(2)
'レコード単位の排他的ロック。編集直後のデータソースでレコードをロックする。

'adLockOptimistic(3)
'レコード単位の共有的ロック。
'Updateメソッドを呼び出した場合にのみ、レコードをロックする。

'adLockBatchOptimistic(4)
'共有的バッチ更新。バッチ更新モードの場合にのみ指定可能。

'adLockUnspecified(-1)
'ロックの種類を指定しない。

'レコード単位のロックが可能
'(ADOの場合は)読み取り専用で開けること
'↑の仕組みを理解しておくことが重要

'==================================================
'【DAO】トランザクション処理
'途中でエラーになった場合の不整合をなくす処理(DAOにもADOにもある)
'==================================================
Private Sub DAOTransactionSample()
#Const ARI = True

#If ARI Then
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
#Else
    Dim ws As Object
    Dim db As Object
    Dim rs As Object
#End If

    Set ws = DAO.DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase("C:\Users\USER\Desktop\Git管理下\CooperationAccessLesson" & "\売上データ.accdb")
    
    'テーブル「T_商品マスタ」をレコードセットして開く
    Set rs = db.OpenRecordset("T_商品マスタ", dbOpenDynaset)

    'エラー処理を開始する
    On Error GoTo ErrHdl
    
    'トランザクション処理を開始---------------------------------------------------
    ws.BeginTrans
    
    rs.AddNew
    'rs!ID = "B0004"
    rs!商品名 = "ベルトD"
    rs!単価 = 12000
    rs.Update
    
    ws.CommitTrans
    '処理を確定---------------------------------------------------
    
ExitHdl:
    rs.Close
    db.Close
    ws.Close
    Exit Sub

ErrHdl:
    ws.Rollback
End Sub

'==================================================
'【ADO】トランザクション処理
'途中でエラーになった場合の不整合をなくす処理(DAOにもADOにもある)
'==================================================
Private Sub ADOTransactionSample()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'Connectiionオブジェクトを作成
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & "C:\Users\USER\Desktop\Git管理下\CooperationAccessLesson" & "\売上データ.accdb;"
    
    rs.Open "T_商品マスタ", cn, adOpenDynamic, adLockOptimistic
    
    'エラー処理開始
    On Error GoTo ErrHdl
    
    cn.BeginTrans '---------------------------------------------------
    rs.AddNew
    'rs!ID =""
    rs!商品名 = "パンツA"
    rs!単価 = 15000
    rs.Update
    cn.CommitTrans '---------------------------------------------------
    
ExitHdl:
    rs.Close
    cn.Close
    Exit Sub
    
ErrHdl:
    cn.RollbackTrans
End Sub

'このあたりのエラー処理、トランザクションはセットで使用する(使えるようになっておくほうがベスト)

'---------------------------------------------------
'Accessそのものの処理を自動化するには
'参照設定【Microsoft Access XX.X Object Library】
'DoCmdオブジェクトを使用する
'Accessフォームを開く、閉じる、クエリを実行するといった操作はできても、
'データを直接更新したりすることはできない
'---------------------------------------------------















































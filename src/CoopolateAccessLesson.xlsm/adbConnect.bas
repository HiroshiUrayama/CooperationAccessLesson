Attribute VB_Name = "adbConnect"
'#################################################
'ADOでAccessに接続する
'#################################################
Option Explicit

'==================================================
'ADOを利用してデータベースに接続するコード
'まさかの、Module名称がADOとか書いてあると参照設定ができん(すでに使われています表示)
'==================================================
Private Sub dbSample()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'Connectionオブジェクトを作成する
    Set cn = New ADODB.Connection
    
    'Recordsetオブジェクトを作成する
    Set rs = New ADODB.Recordset
    
    '「顧客データ.accdb」データベースを開く
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\顧客データ.accdb;"
                
    '「T顧客リスト」テーブルのデータを取得する(adOpenForwardOnly,adLockReaOnlyが既定値)
    rs.Open "T顧客リスト", cn, adOpenForwardOnly, adLockReadOnly
    
    'メッセージを取得する
    MsgBox "【1件目のデータ】" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(4).Name & ":" & rs.Fields(1).Value & vbCrLf
        
    rs.Close
    cn.Close

End Sub

'---------------------------------------------------
'レコードセットの開き方が大切。
'---------------------------------------------------
'object.Open Source, ActiveConnection, CursorType, LockType, Option
'オブジェクトにcn(Connectionオブジェクト)を指定すると、データベースに接続することができる。

'Source：省略可能。テーブル名やクエリ名、SQL分を指定する。
'ActiveConnection：省略可能。Connectionオブジェクトまたは接続情報文字列を指定する。
'CursorType：省略可能。カーソルタイプを決めるための値を指定する。
'LockType：省略可能。ロックの種類を決めるための値を指定する。
'Options：省略可能。

'---------------------------------------------------
'開き方が大切。
'---------------------------------------------------
'★CursorType
'adOpenForwardOnly(0)／(前方専用カーソル)
    '既定値。レコードのスクロール方向が前方向に限定されていることを除き、静的カーソルと同じ働きをする
'adOpenKeyset(1)／(キーセットカーソル)
    '他のユーザが追加したレコードは表示できない。それ以外は動的カーソルと同じ。
'adOpenDynamic(2)／(動的カーソル)
    '他のユーザによる追加、変更、及び削除を確認できる。
    'プロバイダがブックマークをサポートしている場合、Recordset内での全ての操作を許可する
'adOpenStatic(3)／データの検索やレポートの作成に使用するための、レコードの静的コピー。
    '他のユーザーによる追加、変更、削除は表示されない。
'adOpenUnspecified(-1)
    'カーソルの種類を指定しない
    
'★LockType
'adLockReadOnly(1)／既定値。読み取り専用。データの変更不可
'adLockPessimistic(2)／レコード単位の排他的ロック。
                                '編集直後のデータソースでレコードをロックする。
'adLockOptimistic(3)／レコード単位の共有的ロック。Updateメソッドを呼び出した場合にのみレコードをロック。
'adLockBatchOptimistic(4)／共有的バッチ更新。
                                'バッチ更新モードの場合にのみ指定可能。
'adLockUnspecified(-1)／ロックの種類を指定しない


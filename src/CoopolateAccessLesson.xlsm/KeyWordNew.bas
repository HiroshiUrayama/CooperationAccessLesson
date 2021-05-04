Attribute VB_Name = "KeyWordNew"
'#################################################
'Newキーワードのお勉強
'#################################################
Option Explicit

'==================================================
'同時にNewすると何が問題なのか(とりあえずどちらも正しく動く)
'==================================================
Private Sub ADOSample4()
    
#Const SAME = True

#If SAME Then
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
#Else
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#End If
    
    
    'Trueの場合、ここの時点ではまだオブジェクト自体が生成されていない
    'FALSEの場合は、ここの時点ですでにオブジェクトが生成されている
    Stop
    
    '処理自体は一緒なので、変数の宣言部分だけ調節する
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\顧客データ.accdb"
                
    rs.Open "T顧客リスト", cn, adOpenForwardOnly, adLockReadOnly
    
    MsgBox "【1件目のデータ】" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf
        
    Stop
    
    rs.Close
    cn.Close
End Sub

'==================================================
'同時にNewするとまずいパターン
'==================================================
Private Sub NewKeywordSample1()
    
#Const ARI = True

#If ARI Then
    '【ローカルウインドウで確認すれば一発】
    'この書き方は、変数の宣言時にオブジェクトを生成するわけではない。
    '使用する際に「なければ作る」という動きをする(これ以降も同じような動きをする)
    '変数を使用するたびに変数を評価する処理が入るから、わずかだけど遅くなる。
    Dim obj As New Collection
#Else
    Dim obj As Object
    Set obj = New Collection
#End If
    
    '要素を追加する
    obj.Add "A"
    
    'Collectionオブジェクトへの参照を解除する
    '【ARIがTrueのとき】これが間違いなのか、その後のAddが間違いないのかがわかんなくなる
    Set obj = Nothing
    
    '再度要素を追加
    'ARI=Trueのときは勝手にobjが作り出されてBがaddされる(Nothingされているのに再度生成されてしまう)
    'ARI=FALSEのときは勝手にobjが作り出されず、ここできちんとエラーを吐いて止まる(Nothingされた結果が正)
    obj.Add "B"
    Stop
End Sub

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

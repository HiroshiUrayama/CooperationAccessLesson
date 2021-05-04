Attribute VB_Name = "binding"
'#################################################
'事前バインディングと実行時バインディングの違い(Code)
'#################################################
Option Explicit

'==================================================
'事前バインディング
'==================================================
Private Sub sample1()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '「顧客データ.accdb」データベースを開く
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\顧客データ.accdb;"
    
    'レコードセットを取り出す
    rs.Open "T顧客リスト", cn, adOpenForwardOnly, adLockReadOnly
    
    'メッセージ表示
    MsgBox "【1件目のデータ】" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

    rs.Close
    cn.Close
End Sub

'==================================================
'実行時バインディング
'==================================================
Private Sub sample2()
    
    '変数宣言がObject型になっている
    Dim cn As Object
    Dim rs As Object
    
    'CreateObject関数を使用
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    '「顧客データ.accdb」データベースを開く
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\顧客データ.accdb;"
    
    'レコードセットを取り出す
    rs.Open "T顧客リスト", cn, adOpenForwardOnly, adLockReadOnly
    
    'メッセージ表示
    MsgBox "【1件目のデータ】" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

    rs.Close
    cn.Close
End Sub

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'事前バインディングと実行時バインディングのメリット、デメリット
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'バインディング             ’メリット                              ’デメリット
'事前バインディング         ’処理速度が若干早い            ’汎用性にかけるケースが有る
                                'インテリセンスが利用できる
'実行時バインディング       ’処理速度が比較的遅い          ’汎用性がある
                                'インテリセンスが使えない

'※つまり、開発するときは事前バインディングと実行時バインディングのいいとこ取りをすればいいわけで。
'それをどうやって作るのかを考えるのが大切。

'---------------------------------------------------
'汎用性とは：作成したプログラムをを利用する際に、別のPCでもきちんと動作するか？
'---------------------------------------------------
'社内で使用する際に、人によって使用するOfficeのバージョンが違ったりすると、参照設定が使えなくなる
'参照不可になったら、参照設定を調節したらいいけど、VBAを使っていない人は参照設定を知らない人が多いので、手動でやるのが難しい
'CreateObject関数でやると今のバージョンに合わせてくれるので、これが「汎用性が高い」という意味。

'VBAで参照設定の有無を調べられるが、「セキュリティセンター」で「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れる必要がある
'参照設定を解除するメソッドはあるが(Remove)、参照不可のライブラリに対してRemoveメソッドを実行すると、エラーになる
'∴参照不可の場合は手動で解除するしか無い。

'==================================================
'【最終形】
'事前バインディングと実行時バインディングのいいところを取った作業の仕方
'開発するときは事前バインディングで開発
'動作させるときは実行時バインディングで動作させる
'∴開発時にはインテリセンスを使ってさっさと開発して、リリース時には実行時バインディングで汎用性をもたせて開発する
    'リリース時には参照設定を外す手続きは必要。
'==================================================
Private Sub ADOSample3()
    
'条件付きコンパイル定数(実行環境で変更したりする)
#Const HAS_REF = True

'条件付きコンパイルの記号を書いておかないと、コンパイルを全行に亘って行うのでエラーになるはず
#If HAS_REF Then
    '参照設定がONのときはこちらが処理される(つまり、開発環境の時)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#Else
    '参照設定がOFFのときはこちらが処理される(つまり、リリースの時)
    Dim cn As Object
    Dim rs As Object
    cn = CreateObject("ADODB.Connection")
    rs = CreateObject("ADODB.Recordset")
#End If
    
    '処理自体は一緒なので、変数の宣言部分だけ調節する
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\顧客データ.accdb"
                
    rs.Open "T顧客リスト", cn, adOpenForwardOnly, adLockReadOnly
    
    MsgBox "【1件目のデータ】" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

End Sub

'---------------------------------------------------
'ディレクティブ(条件付きコンパイル)
'---------------------------------------------------
'「#Const ディレクティブ」で定数を宣言し、「#If...Then...#Else ディレクティブ」で処理を分ける
'#がついてる部分は分岐内しかコンパイルしない(#定数がTrueであれば、#If文がTrueの部分のみしかコンパイルしない)
'∴True、Falseの時に変数の名前が同じことを書いてあっても大丈夫


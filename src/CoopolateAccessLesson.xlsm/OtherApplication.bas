Attribute VB_Name = "OtherApplication"
'#################################################
'他のアプリケーションとの連携
'#################################################
Option Explicit

'==================================================
'Outlook処理自動化_押さえておくべきはオブジェクトモデル
'==================================================
'outlookはメールだけでなく、スケジュール管理もできるから

'Outlookのオブジェクトモデル
'---------------------------------------------------
'【Applicationオブジェクト】
'Outlookそのものを表す。ただし、Excel含め他のOfficeアプリケーションと異なり、複数のAplicatonオブジェクトを作成できない。
'そのため、既にOutlookが起動している状態で、新たにApplicationオブジェクトを作成しても、
'別のOutlookが起動するんじゃなくて、起動中のOutlookを参照する。

'なお、Excelでは、ExcelのWorkbookオブジェクトを新規作成すると、自動的にアプリケーションが起動する(明示的にApplicationオブジェクトを参照しなくても)。
'しかし、Outlookは最初にApplicationオブジェクトを生成する必要がある。

'【Namespaceオブジェクト】
'アドレス帳やメッセージなどのデータにアクセスするためのインターフェース。
'OutlookをVBAで扱う場合の入り口。
'NameSpaceオブジェクトは、ApplicationオブジェクトのGetNameSpaceメソッドで引数に"MAPI"を指定して取得する。
'なお、GetNameSpaceメソッドの引数に指定できる値は"MAPI"のみ。

'★MAPI
'WindowsでMicrosoft者のアプリケーションから電子メールを扱うための標準仕様のことを言う。

'【Explorerオブジェクト】
'フォルダを表示しているウインドウに対応するオブジェクト。
'フォルダのビューやどのアイテムを選択しているかと言った情報などを取得・設定する。
'ApplicationのExplorersコレクションまたはActiveExplorerプロパティで、Explorerオブジェクト取得できる。

'【Inspectorオブジェクト】
'アイテムの表示を行っているウインドウに対応するオブジェクト。
'表示中のアイテムやアイテムウインドウのコマンドボタン等の情報を取得・設定する。
'ApplicationオブジェクトのInspectorsコレクションまたはActiveInspectorプロパティによって、Inspectorオブジェクトを取得することができる。

'【Folderオブジェクト】
'フォルダに対応するオブジェクト。このオブジェクトを通じて、フォルダ内のメッセージやサブフォルダ、ビューなどの情報にアクセスする。
'受信トレイや送信済みアイテムなど、
'Outlook起動時に作成される規定のフォルダについては、
'NameSpaceオブジェクトのGetDefaultFolderメソッドを使用して取得可能。

'【Itemオブジェクト】
'フォルダに格納されているアイテムには、それぞれのアイテムの種別に応じたオブジェクトがある。
'例えば、メッセージのアイテムはMailItemオブジェクト、予定はAppointmentItemオブジェクト、仕事はTaskItemオブジェクトとなる。
'Itemオブジェクトは、FolderオブジェクトのItemsコレクションから取得できる。
'また、ApplicationオブジェクトのCreateItemメソッドで新規アイテムを作成することも可能。

'Outlookに直接書いたVBAの処理
'---------------------------------------------------
'Excelみたいに配布して使うことができない。
    '∴モジュールを配布して使用する人のOutlookにCodeをコピーする
    'Moduleをインポートする


'==================================================
'新規メールを作成する
'==================================================
Private Sub CreateNewMail()
    Dim olApp As Outlook.Application
    Dim MailItem As Outlook.MailItem
    
    'Outlookを起動する
    Set olApp = New Outlook.Application
        
    'メールを作成する
    Set MailItem = olApp.CreateItem(olMailItem)
    
    With MailItem
        '送信先を指定する
        .Recipients.Add("hiroshiurayama0308@gmail.com").Type = 1
        .Subject = "VBAから送信"
        
        '本文の設定
        .Body = "自分へ" & vbCrLf & _
                    "こんにちは。自分です。" & vbCrLf & _
                    "練習で自分に送信しています。" & vbCrLf & _
                    "今日は生憎の雨ですが、体調はいかがでしょうか?" & vbCrLf & _
                    "昨日からボロネーゼばかりで体調が悪くなりそうなので、" & vbCrLf & _
                    "今日はなめこのお味噌汁とか、脂質の少ないものを食べて回復しましょうね。" & vbCrLf & _
                    "以上です。"
                    
        '添付ファイルを指定する
        .Attachments.Add "C:\Users\USER\Desktop\LineUp用_写真、動画\写真\IMG_20210213_152036.jpg"
        '.Send
        .Display
        
        .Send
    End With
End Sub

'==================================================
'メールの情報を取得するサンプル
'==================================================
Private Sub GetMailItem()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.Folder
    Dim i As Long
    
    'outlookを起動する
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    '対象フォルダを選択する
    Set olFolder = olNamespace.PickFolder
    
    '全てのアイテムに対して処理を行う
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






































































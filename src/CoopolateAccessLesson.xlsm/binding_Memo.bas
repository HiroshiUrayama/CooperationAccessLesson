Attribute VB_Name = "binding_Memo"
'#################################################
'バインディング
'定義：公開されている他のアプリケーションのオブジェクトやメソッド、プロパティを使用するために、そのライブラリを参照すること。
'ライブラリ
'定義：汎用性の高い複数のプログラムを再利用可能な形でひとまとまりにしたもの」
'タイプライブラリ：VBAと他のアプリケーションをつなぐインタフェースになるもの(DAO、ADO、InternetExplorer、Word、検索機能とか、単体では意味をなさないものも対象)
'#################################################
Option Explicit

'==================================================
'公開されている: DAOやADOが公開しているオブジェクトやメソッド､プロパティを使用するということ｡
'==================================================
'自動化したい処理を行うアプリケーションがそのオブジェクトを公開(他のアプリケーションから使えるように)していることで、
'ExcelVBAからそのアプリケーションを操作することができる(Word、PowerPoint、Outlook、AutoCADとかも公開されているから使える)
'逆に、公開されていなければ使えない。

'==================================================
'ライブラリは、自分で作ることもできる。
'VBAではできない。C言語や.net等の言語を利用することでライブラリを作って使用することも可能。
'.netでVBAのライブラリを作成する方法
'https://www.atmarkit.co.jp/fdotnet/dotnettips/1063vbausedotnet/vbausedotnet.html
'==================================================

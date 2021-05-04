Attribute VB_Name = "ObjectCreateMemo"
'#################################################
'CreateObject関数に入れるクラス名
'#################################################
Option Explicit

'CreateObject(class,[servername])
'classには、「appname.objecttype」という構文で作成するオブジェクトのアプリケーション名とクラスを記載する
'servernameは省略されるか空文字列""の場合にはローカルコンピュータが使われる。
'---------------------------------------------------
'CreateObject関数に使われるProgIDは、
'レジストリ HKEY_LOCAL_MACHINE\SOFTWARE\Classesに保存されている。
'---------------------------------------------------

Sub createObjectStates()
    '参照設定Microsoft.Access XX.X Object Library
    Dim a As Object
    Set a = CreateObject("Access.Application")
    
    'DAOのデータベースオブジェクト
    'accdbファイルの場合はMicrosoft Office XX.X Access Database Engine Object Library
    'mdbファイルの場合はMicrosoft DAO 3.6 Object Library
    '64bitバージョンのAccessでは参照不可能
    Dim b As Object
    Set b = CreateObject("DAO.Database")
    
    'ADOの接続文字列(Microsoft ActiveX Data Objects X.X Library)
    Dim c As Object
    Set c = CreateObject("ADODB.Connection")
    
    'ADOのレコードセット(Microsoft ActiveX DataObjects X.X Library)
    Dim d As Object
    Set d = CreateObject("ADODB.Recordset")
    
    'ADOのストリームオブジェクト(Microsoft ActiveX Data Objects X.X Library)
    Dim e As Object
    Set e = CreateObject("ADODB.Stream")
    
    'ADOXのカタログオブジェクト。
    'データベースに接続する時に指定(Microsoft ADOExt. X.X for DDLand Security)
    Dim f As Object
    Set f = CreateObject("ADOX.Catalog")
    
    'Excelのアプリケーション(Microsoft Excel XX.X Object Library)
    Dim g As Object
    Set g = CreateObject("Excel.Application")
    
    'Excelのワークシート(”Microsoft Excel XX.X Object Library)
    Dim h As Object
    Set h = CreateObject("Excel.Worksheet")
    
    'Outlookのアプリケーション(Microsoft Excel XX.X Object Library)
    Dim i As Object
    Set i = CreateObject("Outlook.Application")
    
    'PowerPointのアプリケーション(Microsoft PowerPoint XX.X Object Library)
    Dim j As Object
    Set j = CreateObject("PowerPoint.Application")
    
    'Wordのアプリケーション(Microsoft Word XX.X Object Library)
    Dim k As Object
    Set k = CreateObject("Word.Application")
    
    'InternetExplorer(Microsoft Internet Controls)
    Dim l As Object
    Set l = CreateObject("ADODB.Stream")
    
    'ファイルシステムオブジェクト(Microsoft Scripting Runtime)
    Dim m As Object
    Set m = CreateObject("Scripting.FileSystemObject")
    
    'ディクショナリオブジェクト("Microsoft Scripting Runtime")
    Dim n As Object
    Set n = CreateObject("Scripting.Dictionary")
    
    'WindowsScriptHostのシェルオブジェクト(Windows Script Host Object Model)
    Dim o As Object
    Set o = CreateObject("WScript.Shell")
    
    'Windowsのシェルオブジェクト(Microsoft Shell Controls And Automation)
    Dim p As Object
    Set p = CreateObject("Shell.Application")
    
    '正規表現で使用するオブジェクト(Microsoft VBScript Regular Expressions 5.5)
    Dim q As Object
    Set q = CreateObject("VBScript.RegExp")
    
End Sub

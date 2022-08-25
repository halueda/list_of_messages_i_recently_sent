Option Explicit

' Usage: $0 [outFilename]
'

Const olFolderCalendar = 9
Const olFolderIndex = 6
Const olFolderManagedEmail = 29
Const olMeetingReceivedAndCanceled = 7


'    Set colItems = Session.GetDefaultFolder(olFolderInbox).Items
'    colItems.Restrict "[受信日時] >= #" & dtStart & "# AND " & _
'        "[受信日時] <  #" & dtEnd & "#"

'テスト項目、変なフォルダーのも引っかかるか。subfolderもひっかかるか

'コマンドライン引数（パラメータ）の取得

Dim oParam
Set oParam = WScript.Arguments

'Dim idx
'For idx = 0 To oParam.Count - 1 
'  WScript.echo oParam(idx)
'Next


' 第一引数（あれば）は、書き出しファイル名。 標準出力を意味する - も指定可能
Dim strFileName
strFileName = "D:\tmp\sche.txt"

Dim stmCSVFile 	'As TextStream

If oParam.Count > 0 Then
	strFileName = oParam(0)
End If

If strFileName = "-" Then
    Set stmCSVFile = WScript.StdOut

    'スクリプト・ホストのファイル名を取得
    Dim strHostName
    strHostName = LCase(Mid(WScript.FullName,  InStrRev(WScript.FullName,"\") + 1))

    'もしホストがwscript.exeなら
    If strHostName = "wscript.exe" Then
	WScript.Echo "Usage: cscript.exe $0 [file]; file can be '-' only when executed from cscript.exe"
	WScript.Quit()
    End if

Else
    Dim objFSO 		'As FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set stmCSVFile = objFSO.CreateTextFile(strFileName, True, True)
End If

'stmCSVFile.WriteLine days


' Main を呼び出す
Main stmCSVFile
stmCSVFile.Close


'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
'  Main routine
'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
' Outlookオブジェクトを作って、処理を開始

Public Sub Main(outStream)
    Dim OTApp   'As Outlook.Application
'    If Process.GetProcessesByName("OUTLOOK").Count() > 0 Then
'        ' If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
'         Set OTApp = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application)
'    Else
        ' If not, create a new instance of Outlook and log on to the default profile.
    Set OTApp = CreateObject("Outlook.Application")
'        Dim ns  'As Outlook.NameSpace
'        Set ns = OTApp.GetNamespace("MAPI")
'        ns.Logon "", "", Missing.Value, Missing.Value
'        Set ns = Nothing
'    End If

'   outStream.WriteLine days

    ShortSentMessageDaysAP OTApp, outStream
End Sub

'''''''''''''''
' Calendar オブジェクトを取得して処理を開始
'''''''''''''''

' VBAエディタでは、 以下をしてから進めると良い
'   Set ap=Application

Public Sub ShortSentMessageDaysAP(ap, outStream)
    Dim myNamespace     'As Outlook.NameSpace
    
    Set myNamespace = ap.GetNamespace("MAPI")
    
'    Dim myFolder   'As Outlook.Folder
'
'    Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox)

    Dim colStores   'As Outlook.Stores
'    Set colStores = ap.Session.Stores
    Set colStores = myNamespace.Stores

    Dim oSearchFolders  'As Outlook.Folders
    Set oSearchFolders = colStores("ueda.haruyasu@fujitsu-general.com").GetSearchFolders

    Dim oFolder 'As Outlook.Folder
'    Set oFolder = oSearchFolders("最近自分が出したメール")
'    outStream.WriteLine oSearchFolders(1).Name
'    outStream.WriteLine oSearchFolders(8).Name
    Set oFolder = oSearchFolders("sentmail")
   
    ShortSentMessageDays oFolder, outStream

End Sub


'''''''''''''''
' 処理の本体
'''''''''''''''
' 指定の日数（正の時は未来、負のときは過去）のスケジュールを取得
' MeetingReceivedAndCanceled はスキップ
' 時刻が0000はスキップ
'
' 日付が変わったら日付を出力
' 開始時刻-終了時刻 タイトル@場所 を出力
' ただし、場所は、冗長なテキストを削除
' 最後に出力先をclose
'

Public Sub ShortSentMessageDays(oFolder, outStream)
    Dim strLine     'As String
    Dim mes

    '
    For Each mes In oFolder.Items
        ' if mes.Class <> olMail
        If mes.MessageClass <> "IPM.Note.SMIME" Then
	   strLine = messageLine(mes)
	   outStream.WriteLine strLine
	End If
    Next
End Sub

Function messageLine(mes)
    Dim strLine     'As String

'    日付時刻 Subj12文字(半角24文字)以内 本文54文字

    strLine = Formatmmdd(mes.ReceivedTime)
'    WScript.echo strLine
    strLine = strLine & Formathhnn(mes.ReceivedTime)
'    WScript.echo strLine
    strLine = strLine & " "
'    WScript.echo mes.Subject
    strLine = strLine & MyLeftB(mes.Subject, 48)
'    WScript.echo strLine
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Global = True
        .Pattern = "\t"
        shortBody = .Replace(strLine, "")
        .Pattern = "　*" ' 全角空白は無条件で除く
        shortBody = .Replace(shortBody, "")
    End With

'    On Error GoTo ErrMessageLineHandler

    Dim shortBody

'    body の head 30行を取り出す。
'    body の挨拶、名乗りを除く（最初から五行以内の）お疲れ様です。（いつも）（大変に）お世話に（なっております｜なります）。^.*上田です。
'    body の空行、改行を除く

'    WScript.echo mes.Body
    shortBody = MyLeftB(mes.Body, 200)
'    WScript.echo shortBody

    With RE
        
        .Pattern = "お疲れ様です"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "いつも"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "大変に"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "お世話になっております。"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "お世話になります。"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "-----*"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "=====*"
        shortBody = .Replace(shortBody, "")
        
' 上田専用
        .Pattern = ".*上田です。"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "以下の承認依頼が事務処理支援サービス(InfoShare)にあります"
        shortBody = .Replace(shortBody, "")
        
' 改行や空白を除去
        .Global = True
        
        .Pattern = "\n"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "\r"
        shortBody = .Replace(shortBody, "")
        
'        .Pattern = " *"
'        shortBody = .Replace(shortBody, "")

        .Pattern = "　*" ' 全角空白は無条件で除く
        shortBody = .Replace(shortBody, "")

' 前後どちらかが英数字でないなら空白を除く
        .Pattern = "([^0-z]) *"
        shortBody = .Replace(shortBody, "$1")
        
        .Pattern = " *([^0-z])"
        shortBody = .Replace(shortBody, "$1")
        
        .Pattern = "\t"
        shortBody = .Replace(shortBody, "")
                
    End With
    
'    WScript.echo shortBody
    If shortBody <> "" Then
        strLine = strLine & "|"
        strLine = strLine & shortBody
    End If

    ' shortBody =~ s/.* (.*)会議.*/$1/
'    messageLine =  leftB(strLine, 145)
'ErrMessageLineHandler:
'    On Error GoTo 0
	
'    WScript.echo strLine
    messageLine = MyLeftB(strLine, 160)
'    WScript.echo strLine
End Function

'''''''''''''''
' 時刻を 2359のような 数字4文字に変換
'''''''''''''''
Private Function Formathhnn(d)
    Dim hh, nn
    hh = Right("0" & Hour(d), 2)
    nn = Right("0" & Minute(d), 2)
    Formathhnn = hh & nn '  Format(d, "hhnn")
End Function

'''''''''''''''
' 日付を 04/01火 のような文字列に変換
'''''''''''''''
Private Function Formatmmdd(d)
    Dim mm, dd, aaa
    mm = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d), 2)
    aaa = Mid("日月火水木金土", Weekday(d), 1)
    Formatmmdd = mm & "/" & dd & aaa ' Format(d, "mm/ddaaa")

End Function

''**引数説明*******************************
 ''* pS_String:元の対象文字列
''* pI_Len 　:先頭から抜き出すバイト数
''*****************************************
'Public Function MyLeftB(pS_String As String, pI_Len As Integer) As String
Public Function MyLeftB(pS_String, pI_Len )
 Dim S_Wkstring 'As String '作業用文字列エリア

'On Error GoTo ErrHandler

 ''SJISに変換し、LeftB関数を行う
 Dim bobj
 Set bobj = CreateObject("basp21")
' S_Wkstring = LeftB(pS_String, pI_Len)
' MyLeftB = S_Wkstring

 Dim kcode
 kcode = bobj.KConv(pS_String, 0)
 MyLeftB = pS_String
 ' WScript.echo kcode  ' UNICODE UCS2 = 4
 MyLeftB = bobj.KConv(MyLeftB, 1, kcode)
 MyLeftB = bobj.MidB(MyLeftB, 0, pI_Len)
 MyLeftB = bobj.KConv(MyLeftB, kcode, 1)
 
' S_Wkstring = LeftB(bobj.KConv(pS_String, 1), pI_Len)
' S_Wkstring = bobj.MidB(pS_String, 1, pI_Len)

'  MyLeftB = bobj.KConv(S_Wkstring, kcode, 1)
'  MyLeftB = CStr(bobj.KConv(S_Wkstring, kcode, 1))
 
'S_Wkstring = LeftB(StrConv(pS_String, vbFromUnicode), pI_Len)
' MyLeftB = StrConv(S_Wkstring, vbUnicode)

' On Error GoTo 0
' Exit Function

'ErrHandler:
' MyLeftB = ""
' On Error GoTo 0
End Function


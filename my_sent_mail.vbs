Option Explicit

' Usage: $0 [outFilename]
'

Const olFolderCalendar = 9
Const olFolderIndex = 6
Const olFolderManagedEmail = 29
Const olMeetingReceivedAndCanceled = 7


'    Set colItems = Session.GetDefaultFolder(olFolderInbox).Items
'    colItems.Restrict "[��M����] >= #" & dtStart & "# AND " & _
'        "[��M����] <  #" & dtEnd & "#"

'�e�X�g���ځA�ςȃt�H���_�[�̂����������邩�Bsubfolder���Ђ������邩

'�R�}���h���C�������i�p�����[�^�j�̎擾

Dim oParam
Set oParam = WScript.Arguments

'Dim idx
'For idx = 0 To oParam.Count - 1 
'  WScript.echo oParam(idx)
'Next


' �������i����΁j�́A�����o���t�@�C�����B �W���o�͂��Ӗ����� - ���w��\
Dim strFileName
strFileName = "D:\tmp\sche.txt"

Dim stmCSVFile 	'As TextStream

If oParam.Count > 0 Then
	strFileName = oParam(0)
End If

If strFileName = "-" Then
    Set stmCSVFile = WScript.StdOut

    '�X�N���v�g�E�z�X�g�̃t�@�C�������擾
    Dim strHostName
    strHostName = LCase(Mid(WScript.FullName,  InStrRev(WScript.FullName,"\") + 1))

    '�����z�X�g��wscript.exe�Ȃ�
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


' Main ���Ăяo��
Main stmCSVFile
stmCSVFile.Close


'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
'  Main routine
'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
' Outlook�I�u�W�F�N�g������āA�������J�n

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
' Calendar �I�u�W�F�N�g���擾���ď������J�n
'''''''''''''''

' VBA�G�f�B�^�ł́A �ȉ������Ă���i�߂�Ɨǂ�
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
'    Set oFolder = oSearchFolders("�ŋߎ������o�������[��")
'    outStream.WriteLine oSearchFolders(1).Name
'    outStream.WriteLine oSearchFolders(8).Name
    Set oFolder = oSearchFolders("sentmail")
   
    ShortSentMessageDays oFolder, outStream

End Sub


'''''''''''''''
' �����̖{��
'''''''''''''''
' �w��̓����i���̎��͖����A���̂Ƃ��͉ߋ��j�̃X�P�W���[�����擾
' MeetingReceivedAndCanceled �̓X�L�b�v
' ������0000�̓X�L�b�v
'
' ���t���ς��������t���o��
' �J�n����-�I������ �^�C�g��@�ꏊ ���o��
' �������A�ꏊ�́A�璷�ȃe�L�X�g���폜
' �Ō�ɏo�͐��close
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

'    ���t���� Subj12����(���p24����)�ȓ� �{��54����

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
        .Pattern = "�@*" ' �S�p�󔒂͖������ŏ���
        shortBody = .Replace(shortBody, "")
    End With

'    On Error GoTo ErrMessageLineHandler

    Dim shortBody

'    body �� head 30�s�����o���B
'    body �̈��A�A�����������i�ŏ�����܍s�ȓ��́j�����l�ł��B�i�����j�i��ςɁj�����b�Ɂi�Ȃ��Ă���܂��b�Ȃ�܂��j�B^.*��c�ł��B
'    body �̋�s�A���s������

'    WScript.echo mes.Body
    shortBody = MyLeftB(mes.Body, 200)
'    WScript.echo shortBody

    With RE
        
        .Pattern = "�����l�ł�"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "����"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "��ς�"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "�����b�ɂȂ��Ă���܂��B"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "�����b�ɂȂ�܂��B"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "-----*"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "=====*"
        shortBody = .Replace(shortBody, "")
        
' ��c��p
        .Pattern = ".*��c�ł��B"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "�ȉ��̏��F�˗������������x���T�[�r�X(InfoShare)�ɂ���܂�"
        shortBody = .Replace(shortBody, "")
        
' ���s��󔒂�����
        .Global = True
        
        .Pattern = "\n"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "\r"
        shortBody = .Replace(shortBody, "")
        
'        .Pattern = " *"
'        shortBody = .Replace(shortBody, "")

        .Pattern = "�@*" ' �S�p�󔒂͖������ŏ���
        shortBody = .Replace(shortBody, "")

' �O��ǂ��炩���p�����łȂ��Ȃ�󔒂�����
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

    ' shortBody =~ s/.* (.*)��c.*/$1/
'    messageLine =  leftB(strLine, 145)
'ErrMessageLineHandler:
'    On Error GoTo 0
	
'    WScript.echo strLine
    messageLine = MyLeftB(strLine, 160)
'    WScript.echo strLine
End Function

'''''''''''''''
' ������ 2359�̂悤�� ����4�����ɕϊ�
'''''''''''''''
Private Function Formathhnn(d)
    Dim hh, nn
    hh = Right("0" & Hour(d), 2)
    nn = Right("0" & Minute(d), 2)
    Formathhnn = hh & nn '  Format(d, "hhnn")
End Function

'''''''''''''''
' ���t�� 04/01�� �̂悤�ȕ�����ɕϊ�
'''''''''''''''
Private Function Formatmmdd(d)
    Dim mm, dd, aaa
    mm = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d), 2)
    aaa = Mid("�����ΐ��؋��y", Weekday(d), 1)
    Formatmmdd = mm & "/" & dd & aaa ' Format(d, "mm/ddaaa")

End Function

''**��������*******************************
 ''* pS_String:���̑Ώە�����
''* pI_Len �@:�擪���甲���o���o�C�g��
''*****************************************
'Public Function MyLeftB(pS_String As String, pI_Len As Integer) As String
Public Function MyLeftB(pS_String, pI_Len )
 Dim S_Wkstring 'As String '��Ɨp������G���A

'On Error GoTo ErrHandler

 ''SJIS�ɕϊ����ALeftB�֐����s��
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


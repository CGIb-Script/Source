VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Hash As New MD5Hash
Private bytBlock() As Byte
Public Sub ConsoleColor(ByVal szin As String)
Dim scrbuf As CONSOLE_SCREEN_BUFFER_INFO
If UCase(szin) = "BLACK" Then SetConsoleTextAttribute hStdOut, FOREGROUND_BLACK Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "BLUE" Then SetConsoleTextAttribute hStdOut, FOREGROUND_BLUE Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "GREEN" Then SetConsoleTextAttribute hStdOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "CYAN" Then SetConsoleTextAttribute hStdOut, FOREGROUND_CYAN Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "RED" Then SetConsoleTextAttribute hStdOut, FOREGROUND_RED Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "MAGENTA" Then SetConsoleTextAttribute hStdOut, FOREGROUND_MAGENTA Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "YELLOW" Then SetConsoleTextAttribute hStdOut, FOREGROUND_BROWN Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "WHITE" Then SetConsoleTextAttribute hStdOut, FOREGROUND_LIGHTGRAY Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
If UCase(szin) = "LIGHTGRAY" Then SetConsoleTextAttribute hStdOut, FOREGROUND_INTENSITY Or FOREGROUND_INTENSITY: hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

End Sub
Public Function StrXML(ByVal ifi As String, ByVal searchStr As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As String
Dim jo As String
  
  For t = 1 To Len(ifi)
   s = Mid(ifi, t, Len(searchStr))
   If s = searchStr Then
            For h = t To Len(ifi)
               so = Mid(ifi, h, 2)
               If so = "</" Then jo = Mid(ifi, t + Len(searchStr) + 1, h - t - Len(searchStr) - 1): Exit For
            Next
          Exit For
     End If
   Next
 
tempVar = jo
StrXML = Trim(tempVar)
End Function

Public Function RunBat(ByVal LogFile As String) As String
 RunBat = Shell(LogFile, vbHide)
End Function


Public Function ping(ByVal ifi As String) As String
On Error Resume Next
 Dim ResolveResult As RESOLVE_ERROR_ENUM
    Dim IP As String
 
    
      
    ResolveResult = ResolveData(ifi, IP)
    If ResolveResult = RES_SUCCESS Then
          ping = "True," & IP
    Else
        ping = "False"
    End If
    
End Function

Public Function Upload(ByVal tPath As String, ByVal tFileName As String) As String
 GetFormDataRU
 sUploadDir = tPath
 sUploadFilename = tFileName
 sUploadFileNamess = ""

If InStr(1, CGI_ContentType, "multipart", 1) > 0 Then
   MultiPart sFormData
 Else
   fUpload = "File not received. Your browser did not send the expected multipart headers."
   Exit Function
End If

Upload = sUploadFileNamess


End Function
Sub StorePairsRU(sData As String)
Dim pointer    As Long      ' sData position pointer
Dim n          As Long      ' name/value pair counter
Dim delim1     As Long      ' position of "="
Dim delim2     As Long      ' position of "&"
Dim lastPair   As Long      ' size of tPair() array
Dim lPairs     As Long      ' number of name=value pairs in sData

lastPair = UBound(tPair)    ' current size of tPair()
delim1 = 0
Do
  delim1 = InStr(delim1 + 1, sData, "=")
  If delim1 = 0 Then Exit Do
  lPairs = lPairs + 1
Loop

If lPairs = 0 Then Exit Sub  'nothing to add

' redim tPair() based on the number of pairs found in sData
ReDim Preserve tPair(lastPair + lPairs) As pair

' assign values to tPair().name and tPair().value
pointer = 1
For n = (lastPair + 1) To UBound(tPair)
   delim1 = InStr(pointer, sData, "=") ' find next equal sign
   If delim1 = 0 Then Exit For         ' parse complete

   tPair(n).name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))
   
   delim2 = InStr(delim1, sData, "&")

   ' if no trailing ampersand, we are at the end of data
   If delim2 = 0 Then delim2 = Len(sData) + 1
 
   ' value is between the "=" and the "&"
   tPair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
   pointer = delim2 + 1
Next n
End Sub


Sub GetFormDataRU()

Dim sBuff        As String      ' buffer to receive POST method data
Dim lBytesRead   As Long        ' actual bytes read by ReadFile()
Dim rc           As Long        ' return code


'
If CGI_RequestMethod = "POST" Then
   sBuff = String(lContentLength, Chr$(0))
   Do While Len(sFormData) < lContentLength
      rc = ReadFile(hStdIn, ByVal sBuff, lContentLength, lBytesRead, ByVal 0&)
      sFormData = sFormData & Left$(sBuff, lBytesRead)
   Loop
   '
   If InStr(1, CGI_ContentType, "www-form-urlencoded", 1) Then
      StorePairs sFormData
   End If
End If
StorePairsRU CGI_QueryString
End Sub

Public Function Command() As String
Command = Command$
End Function
Public Function MDEncode(ByVal xData As String) As String
bytBlock = StrConv(xData, vbFromUnicode)
    MDEncode = Hash.HashBytes(bytBlock)
End Function

Public Function getVar(ByVal xData As Integer) As String
getVar = vb_pub(xData)
End Function


Public Function Var(ByVal xData As Integer, Optional ByVal xVar As String) As String
If xVar <> "" Then vb_pub(xData) = xVar
Var = vb_pub(xData)
End Function


Public Function HTTPRequest(ByVal URL As String) As String
 Dim http
 Set http = CreateObject("MSXML2.ServerXMLHTTP")
 http.Open "GET", URL, False
 http.Send
 HTTPRequest = http.responseText
 DoEvents
End Function

Public Function RequestURL(ByVal URL As String) As String
Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim sUrl                As String
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

    sUrl = URL
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
       
    RequestURL = sBuffer
    
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    
   
End Function

Public Function Download(ByVal sUrl As String, ByVal xPath As String)
Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

   
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    Open xPath For Binary Access Write As #1
    Put #1, , sBuffer
    Close #1
        
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    
    Download = xPath
End Function
Public Function parseFloat(Data) As Integer
 On Error Resume Next
 eme = Data
 
 eme = Replace(eme, ",", ".")
 
 For g = 0 To 45
  eme = Replace(eme, Chr(g), "")
 Next
 
 eme = Replace(eme, Chr(47), "")
 
 For g = 58 To 255
  eme = Replace(eme, Chr(g), "")
 Next
 
 If eme = "" Then eme = 0
 
 parseFloat = eme
End Function
 


Public Function TimeStamp(ByVal Data As Integer)
If Data = 1 Then
   TimeStamp = (DateDiff("s", 25569#, Now()) * 1000#) - 2202518
   Else
   TimeStamp = (DateDiff("s", 25569#, Now()) * 1000#)
End If


End Function


Public Sub DoEvent()
 DoEvents
End Sub
Public Sub Alert(ByVal nfil As String)
 SendE "<script>alert('" & nfil & "');</script>"
End Sub
Public Function ContentType(Optional ByVal strText As String) As String
 If strText <> "" Then Cont_Type = strText
 ContentType = Cont_Type
End Function
Public Sub Echo(ByVal fileN As String)
 fileN = Replace(fileN, "<cgib>", "&#60;cc&#62;")
 fileN = Replace(fileN, "</cgib>", "&#60;&#47;cc&#62;")
 SendE fileN
End Sub
Public Function MyType() As String
MyType = LCase(KIT)
End Function

Public Sub ScriptInfo()
If KIT <> "CMD" Then
SendE "&nbsp; DOCUMENT_ROOT = " & Environ("DOCUMENT_ROOT") & "<br>"
SendE "&nbsp; HTTP_ACCEPT = " & Environ("HTTP_ACCEPT") & "<br>"
SendE "&nbsp; HTTP_ACCEPT_CHARSET = " & Environ("HTTP_ACCEPT_CHARSET") & "<br>"
SendE "&nbsp; HTTP_ACCEPT_ENCODING = " & Environ("HTTP_ACCEPT_ENCODING") & "<br>"
SendE "&nbsp; HTTP_ACCEPT_LANGUAGE = " & Environ("HTTP_ACCEPT_LANGUAGE") & "<br>"
SendE "&nbsp; HTTP_HOST = " & Environ("HTTP_HOST") & "<br>"
SendE "&nbsp; HTTP_COOKIE = " & Environ("HTTP_COOKIE") & "<br>"
SendE "&nbsp; HTTP_CONNECTION = " & Environ("HTTP_CONNECTION") & "<br>"
SendE "&nbsp; HTTP_REFERER = " & Environ("HTTP_REFERER") & "<br>"
SendE "&nbsp; HTTP_USER_AGENT = " & Environ("HTTP_USER_AGENT") & "<br>"
SendE "&nbsp; sFormData = " & sFormData & "<br>"
SendE "<hr>"
SendE "&nbsp; HTTPS = " & Environ("HTTPS") & "<br>"
SendE "&nbsp; AUTH_TYPE = " & Environ("AUTH_TYPE") & "<br>"
SendE "<hr>"
SendE "&nbsp; CONTENT_LENGTH = " & Environ("CONTENT_LENGTH") & "<br>"
SendE "&nbsp; CONTENT_TYPE = " & Environ("CONTENT_TYPE") & "<br>"
SendE "<hr>"
SendE "&nbsp; GATEWAY_INTERFACE = " & Environ("GATEWAY_INTERFACE") & "<br>"
SendE "<hr>"
SendE "&nbsp; PATH_INFO = " & Environ("PATH_INFO") & "<br>"
SendE "&nbsp; PATH_TRANSLATED = " & Environ("PATH_TRANSLATED") & "<br>"
SendE "<hr>"
SendE "&nbsp; QUERY_STRING = " & Environ("QUERY_STRING") & "<br>"
SendE "<hr>"
SendE "&nbsp; REMOTE_ADDR = " & Environ("REMOTE_ADDR") & "<br>"
SendE "&nbsp; REMOTE_HOST = " & Environ("REMOTE_HOST") & "<br>"
SendE "&nbsp; REMOTE_IDENT = " & Environ("REMOTE_IDENT") & "<br>"
SendE "&nbsp; REMOTE_PORT = " & Environ("REMOTE_PORT") & "<br>"
SendE "&nbsp; REMOTE_USER = " & Environ("REMOTE_USER") & "<br>"
SendE "<hr>"
SendE "&nbsp; REQUEST_METHOD = " & Environ("REQUEST_METHOD") & "<br>"
SendE "&nbsp; REQUEST_TIME = " & Environ("REQUEST_TIME") & "<br>"
SendE "&nbsp; REQUEST_URI = " & Environ("REQUEST_URI") & "<br>"
SendE "<hr>"
SendE "&nbsp; SCRIPT_NAME = " & Environ("SCRIPT_NAME") & "<br>"
SendE "&nbsp; SCRIPT_FILENAME = " & Environ("SCRIPT_FILENAME") & "<br>"
SendE "<hr>"
SendE "&nbsp; SERVER_ADDR = " & Environ("SERVER_ADDR") & "<br>"
SendE "&nbsp; SERVER_ADMIN = " & Environ("SERVER_ADMIN") & "<br>"
SendE "&nbsp; SERVER_NAME = " & Environ("SERVER_NAME") & "<br>"
SendE "&nbsp; SERVER_PORT = " & Environ("SERVER_PORT") & "<br>"
SendE "&nbsp; SERVER_SOFTWARE = " & Environ("SERVER_SOFTWARE") & "<br>"
SendE "&nbsp; SERVER_PROTOCOL = " & Environ("SERVER_PROTOCOL") & "<br>"
SendE "&nbsp; SERVER_SIGNATURE = " & Environ("SERVER_SIGNATURE") & "<br>"
SendE "&nbsp;[END]"
Else
SendE "CGIb Script interpreter" & vbCrLf
SendE "filename: " & cbFile & vbCrLf
SendE "path: " & cbFile & vbCrLf
End If
End Sub
Public Function StrSearch(ByVal ifi As String, ByVal searchStr As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As String
Dim jo As Integer
  jo = 0
  For t = 1 To Len(ifi)
   s = Mid(ifi, t, Len(searchStr))
   If s = searchStr Then jo = jo + 1
   Next
tempVar = Str(jo)
StrSearch = Trim(tempVar)
End Function
Public Function strCropNum(ByVal Data, ByVal kezdo, ByVal vege, ByVal darab)
On Error Resume Next
i = 0
For t = 1 To Len(Data)
 kereses = Mid(Data, t, Len(kezdo))
 If kereses = kezdo Then i = i + 1
 If i = darab Then Pos = t: Exit For
Next
Data2 = Mid(Data, Pos, Len(Data))
strCropNum = kezdo & StrCrop(Data2, kezdo, vege) & vege
End Function
Public Function Request(ByVal este As String) As String
 On Error Resume Next
 Dim tempVar2
 Request = ""
 Request = GetCgiValue(este)
 DoEvent
  DoEvent
  DoEvent
  DoEvent
End Function
Public Function Script(ByVal iFunct As String)
 SendE "<script>" & iFunct & "</script>"
End Function

Public Sub exe(ByVal filep As String)
Dim lik
lik = Shell(filep, 1)
DoEvents
 DoEvents
End Sub
Private Function SelectPrinter(nPrn As String) As Boolean
    Dim prn As Printer
    
    For Each prn In Printers
        If prn.DeviceName = nPrn Then
            SelectPrinter = True
            Set Printer = prn
            Exit Function
        End If
    Next
End Function

Public Function LongText(Optional ByVal Data)
 If Data <> "" Then CGIForm.RequestText = Data
 LongText = CGIForm.RequestText
End Function
Public Function RunScript(ByVal dat As String)
 Dim d As String
 If dat <> "" Then d = d & vbCrLf & dat
 If IncludeType = 0 Then ScriptVB.ExecuteStatement d
 If IncludeType = 1 Then ScriptVB2.ExecuteStatement d
End Function
Public Sub Ent(ByVal num As Integer)
 For t = 1 To Fix(num)
 If KIT <> "CMD" Then SendE "<br/>"
 If KIT = "CMD" Then SendE vbCrLf
 Next
End Sub
Public Sub Forward(ByVal dat As String)
IncludeType = 1
OpenFileForward = dat
AllSite = ""


If Mid(dat, 2, 2) <> ":\" Then dat = vb_appPath & "\" & dat

fileN = FreeFile
If dat <> "" Then
 Open dat For Input As #fileN
  Do While Not EOF(fileN)
   Line Input #fileN, tempVar
   AllSite = AllSite & tempVar & vbNewLine
  Loop
 Close #fileN
End If
    
Scripts = SSearch(AllSite, "<cgib>")

If Fix(Scripts) > 0 Then
 For Interpreter = 1 To Fix(Scripts)
  SC = SCrop(AllSite, "<cgib>", "</cgib>")
  SCL = Replace(SC, "<cgib>", "")
  SCL = Replace(SCL, "</cgib>", "")
  ScriptNow = "": ScriptVB2.AddCode SCL
  AllSite = Replace(AllSite, SC, ScriptNow, 1, Interpreter)
 Next Interpreter
End If

'Send AllSite

End Sub
Public Function StrCrop(ByVal ifi As String, ByVal startStr1 As String, ByVal endStr1 As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As Long
Dim e As Long
Dim h As Long
Dim sS As Long
Dim vV As Long

keresh = Len(startStr1)
s = InStr(1, ifi, startStr1)
e = InStr(s + keresh, ifi, endStr1)
sS = s + keresh
vV = e - (sS)
tempVar = Mid(ifi, sS, vV)
If Err.Description <> "" Then tempVar = ""
StrCrop = tempVar
End Function
Public Function Include(ByVal dat As String)

Dim a As String
Dim d As String
OpenFileForward = dat
If Mid(dat, 2, 2) <> ":\" Then dat = vb_appPath & "\" & dat

f1 = FreeFile
Open dat For Input As #f1
Do While Not EOF(f1)
 Line Input #f1, a
 a = Replace(a, LCase("<cgib>"), "")
 a = Replace(a, LCase("</cgib>"), "")
 If a <> "" Then d = d & vbCrLf & a
Loop
Close #f1

  SourceNowAll = d

If IncludeType = 0 Then ScriptVB.ExecuteStatement d
If IncludeType = 1 Then ScriptVB2.ExecuteStatement d

End Function
Public Function HtmlTextDecode(ByVal strHTML As String) As String
On Error Resume Next
  Dim oDoc As HTMLDocument
  Set oDoc = New HTMLDocument
  oDoc.body.innerHTML = strHTML
  HtmlTextDecode = oDoc.body.innerText
End Function
Public Function Include_HTML(ByVal dat As String)
Dim a As String
Dim d As String

If Mid(dat, 2, 2) <> ":\" Then dat = vb_appPath & "\" & dat

f1 = FreeFile
Open dat For Input As #f1
Do While Not EOF(f1)
 Line Input #f1, a
 If a <> "" Then d = d & vbCrLf & a
Loop
Close #f1
SendE d
End Function
Public Sub Exit_Script()
On Error Resume Next
Close All

If KIT <> "CMD" Then
 Cont_Type = "text/html"
 Send "Status: 200 OK"
 Send "Content-type: " & Cont_Type & vbNewLine
End If

DoEvents
Unload Me
End
End Sub
Public Sub TimeOut_Off()
CGIForm.Timer1.Enabled = False
End Sub
Public Sub TimeOut_On()
CGIForm.Timer1.Enabled = True
End Sub

Public Sub TimeOut_Script(ByVal scriptData As String)
TimeOutScriptData = scriptData
End Sub
Public Function Timeout(ByVal n1 As Integer)
SetTimeIntervall = Fix(n1)
End Function
Sub Wait(ByVal TimeToWait As Long) 'Time In seconds
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait
Do Until GetTickCount > EndTime
DoEvents
Loop
End Sub
Public Function textToID(ByVal Data As String) As String
On Error Resume Next
maid = Data
maid = LCase(maid)
maid = Replace(maid, Chr(34), "")
maid = Replace(maid, "'", "")
maid = Replace(maid, Chr(13), "")
maid = Replace(maid, Chr(10), "")
maid = Replace(maid, "õ", "o")
maid = Replace(maid, "ú", "u")
maid = Replace(maid, "ó", "o")
maid = Replace(maid, "ö", "o")
maid = Replace(maid, "û", "u")
maid = Replace(maid, "é", "e")
maid = Replace(maid, "á", "a")
maid = Replace(maid, "ü", "u")
maid = Replace(maid, "í", "i")


For h = 0 To 31
   maid = Replace(maid, Chr(h), "")
 Next

For h = 33 To 47
   maid = Replace(maid, Chr(h), "")
Next

 For h = 58 To 96
   maid = Replace(maid, Chr(h), "")
 Next

 For h = 123 To 255
   maid = Replace(maid, Chr(h), "")
 Next
 
maid = Replace(maid, "    ", " ")
maid = Replace(maid, "   ", " ")
maid = Replace(maid, "  ", " ")
maid = Replace(maid, " ", "-")
maid = Replace(maid, "----", "-")
maid = Replace(maid, "---", "-")
maid = Replace(maid, "--", "-")

textToID = Mid(maid, 1, 250)
End Function




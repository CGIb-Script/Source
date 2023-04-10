Attribute VB_Name = "CGI4VB"





Public TimeOutScriptData As String
Public PublicFolder As String
Public AllSite As String
Public ScriptNow As String
Public vb_pub(32000) As String
Public vb_App_Path As String
Public ScriptVB As Object
Public ScriptVB2 As Object
Public ScriptVB3 As Object
Public ScriptVBRun As Object
Public ScriptJS As Object
Public IncludeType As Integer
Public eSource As String
Public SourceNowAll As String
Public SetTime As Integer
Public SetTimeIntervall As Integer

Public OpenFileName As String
Public OpenFileForward As String

Declare Function GetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    lpOverlapped As Any) As Long
Declare Function SetFilePointer Lib "kernel32" _
   (ByVal hFile As Long, _
   ByVal lDistanceToMove As Long, _
   lpDistanceToMoveHigh As Long, _
   ByVal dwMoveMethod As Long) As Long
Declare Function SetEndOfFile Lib "kernel32" _
   (ByVal hFile As Long) As Long

'''''''''''''''''''' szin beallitas
Public Type COORD
        x As Integer
        y As Integer
End Type

Public Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type
Public Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type
Public Declare Function GetConsoleScreenBufferInfo Lib "kernel32" _
(ByVal hConsoleOutput As Long, _
lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Public Declare Function SetConsoleTextAttribute Lib "kernel32" _
(ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Public Const FOREGROUND_BLACK = &H0
Public Const FOREGROUND_BLUE = &H1
Public Const FOREGROUND_GREEN = &H2
Public Const FOREGROUND_CYAN = &H3
Public Const FOREGROUND_RED = &H4
Public Const FOREGROUND_MAGENTA = &H5
Public Const FOREGROUND_BROWN = &H6
Public Const FOREGROUND_LIGHTGRAY = &H7
Public Const FOREGROUND_INTENSITY = &H8
Public Const FOREGROUND_WHITE = &H15

' idaig

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const FILE_BEGIN = 0&
Public CGI_Accept            As String
Public CGI_AuthType          As String
Public CGI_ContentLength     As String
Public CGI_ContentType       As String
Public CGI_GatewayInterface  As String
Public CGI_PathInfo          As String
Public CGI_PathTranslated    As String
Public CGI_QueryString       As String
Public CGI_Referer           As String
Public CGI_RemoteAddr        As String
Public CGI_RemoteHost        As String
Public CGI_RemoteIdent       As String
Public CGI_RemoteUser        As String
Public CGI_RequestMethod     As String
Public CGI_ScriptName        As String
Public CGI_ServerSoftware    As String
Public CGI_ServerName        As String
Public CGI_ServerPort        As String
Public CGI_ServerProtocol    As String
Public CGI_UserAgent         As String
Public CGI_sFormData         As String

Public lContentLength As Long
Public hStdIn         As Long
Public hStdOut        As Long
Public sErrorDesc     As String
Public sEmail         As String
Public sFormData      As String

Type pair
  name As String
  Value As String
End Type

Public tPair() As pair



















Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Charrange
  cpMin As Long
  cpMax As Long
End Type

Public Type FormatRange
  hDC As Long
  hdcTarget As Long
  rc As RECT
  rcPage As RECT
  chrg As Charrange
End Type

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public mMargin As Single
Sub Main()
On Error Resume Next

SetTimeIntervall = 3
Load CGIForm

On Error GoTo ErrorRoutine
InitCgi
GetFormData
CGI_Main

EndPgm:
   On Error Resume Next
   End

ErrorRoutine:

   eSource = Replace(Err.Source, "Microsoft", "CGIb")
   eSource = Replace(eSource, "VBScript", "Script")
   If KIT <> "CMD" Then
       sErrorDesc = Err.Description & " Error Number = " & Str$(Err.Number) & vbCrLf & eSource & "<br/>"
       Else
       sErrorDesc = Err.Description & " Error Number = " & Str$(Err.Number) & vbCrLf & eSource & vbCrLf
   End If
   
   ErrorHandler
   Resume EndPgm
End Sub

Sub ErrorHandler()
Dim rc As Long

On Error Resume Next

rc = SetFilePointer(hStdOut, 0&, 0&, FILE_BEGIN)








If Trim(KIT) <> "CMD" Then
 SendHeader "ok"
 Send "<HTML><HEAD><TITLE>Script Error</TITLE></HEAD>"
 Send "<H1>Error in CGIb</H1>"
 Send "The following internal error has occurred:"
 Send "<PRE>" & sErrorDesc & "</PRE>"
 Send "<PRE><font color=#ee0000>" & OpenFileName & " (" & OpenFileForward & ")</font></PRE>"

If IncludeType = 0 Then
 Send "<PRE><b>Description:   </b>" & ScriptVB.Error.Description & "<br/>"
 Send "<b>Script Text:   </b>" & ScriptVB.Error.Text & "<br/>"
 Send "<b>Error Number:  </b>" & ScriptVB.Error.Number & "<br/>"
 Send "<b>Line Number:   </b>" & ScriptVB.Error.Line & "<br/>"
 Send "<b>Column Number: </b>" & ScriptVB.Error.Column & "</PRE><br/>"
End If

If IncludeType = 1 Then
 Send "<PRE><b>Description:   </b>" & ScriptVB2.Error.Description & "<br/>"
 Send "<b>Script Text:   </b>" & ScriptVB2.Error.Text & "<br/>"
 Send "<b>Error Number:  </b>" & ScriptVB2.Error.Number & "<br/>"
 Send "<b>Line Number:   </b>" & ScriptVB2.Error.Line & "<br/>"
 Send "<b>Column Number: </b>" & ScriptVB2.Error.Column & "</PRE><br/>"
End If



SendFooter
End If

If KIT = "CMD" Then
  Send "Syntax Error!"
  Send "The following internal error has occurred:"
  Send sErrorDesc
  If Trim(OpenFileForward) <> "" Then Send "File: " & OpenFileName & " (" & OpenFileForward & ")"

 If IncludeType = 0 Then
  Send "Description: " & ScriptVB.Error.Description
  Send "Script Text: " & ScriptVB.Error.Text
  Send "Error Number: " & ScriptVB.Error.Number
  Send "Line Number: " & ScriptVB.Error.Line
  Send "Column Number: " & ScriptVB.Error.Column
 End If

 If IncludeType = 1 Then
  Send "Description: " & ScriptVB2.Error.Description
  Send "Script Text: " & ScriptVB2.Error.Text
  Send "Error Number: " & ScriptVB2.Error.Number
  Send "Line Number: " & ScriptVB2.Error.Line
  Send "Column Number: " & ScriptVB2.Error.Column
 End If

End If




rc = SetEndOfFile(hStdOut)

End Sub

Sub InitCgi()

hStdIn = GetStdHandle(STD_INPUT_HANDLE)
hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

sEmail = "info@cgib.hu"

CGI_Accept = Environ("HTTP_ACCEPT")
CGI_AuthType = Environ("AUTH_TYPE")
CGI_ContentLength = Environ("CONTENT_LENGTH")
CGI_ContentType = Environ("CONTENT_TYPE")
CGI_GatewayInterface = Environ("GATEWAY_INTERFACE")
CGI_PathInfo = Environ("PATH_INFO")
CGI_PathTranslated = Environ("PATH_TRANSLATED")
CGI_QueryString = Environ("QUERY_STRING")
CGI_Referer = Environ("HTTP_REFERER")
CGI_RemoteAddr = Environ("REMOTE_ADDR")
CGI_RemoteHost = Environ("REMOTE_HOST")
CGI_RemoteIdent = Environ("REMOTE_IDENT")
CGI_RemoteUser = Environ("REMOTE_USER")
CGI_RequestMethod = Environ("REQUEST_METHOD")
CGI_ScriptName = Environ("SCRIPT_NAME")
CGI_ServerSoftware = Environ("SERVER_SOFTWARE")
CGI_ServerName = Environ("SERVER_NAME")
CGI_ServerPort = Environ("SERVER_PORT")
CGI_ServerProtocol = Environ("SERVER_PROTOCOL")
CGI_UserAgent = Environ("HTTP_USER_AGENT")

lContentLength = Val(CGI_ContentLength)
ReDim tPair(0)


End Sub

Sub GetFormData()
Dim sBuff      As String
Dim lBytesRead As Long
Dim rc         As Long

If CGI_RequestMethod = "POST" Then
   sBuff = String(lContentLength, Chr$(0))
   rc = ReadFile(hStdIn, ByVal sBuff, lContentLength, lBytesRead, ByVal 0&)
   sFormData = Left$(sBuff, lBytesRead)
   
   If InStr(1, CGI_ContentType, "www-form-urlencoded", 1) Then
      StorePairs sFormData
   End If
End If

CGI_sFormData = sFormData

StorePairs CGI_QueryString
End Sub

Sub StorePairs(sData As String)
Dim pointer    As Long
Dim n          As Long
Dim delim1     As Long
Dim delim2     As Long
Dim lastPair   As Long
Dim lPairs     As Long

lastPair = UBound(tPair)
pointer = 1
Do
  delim1 = InStr(pointer, sData, "=")
  If delim1 = 0 Then Exit Do
  pointer = delim1 + 1
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

Public Function UrlDecode(ByVal sEncoded As String) As String
Dim pointer    As Long      ' sEncoded position pointer
Dim Pos        As Long      ' position of InStr target

If sEncoded = "" Then Exit Function

' convert "+" to space
pointer = 1
Do
   Pos = InStr(pointer, sEncoded, "+")
   If Pos = 0 Then Exit Do
   Mid$(sEncoded, Pos, 1) = " "
   pointer = Pos + 1
Loop
    
' convert "%xx" to character
pointer = 1

On Error GoTo errorUrlDecode

Do
   Pos = InStr(pointer, sEncoded, "%")
   If Pos = 0 Then Exit Do
   
   Mid$(sEncoded, Pos, 1) = Chr$("&H" & (Mid$(sEncoded, Pos + 1, 2)))
   sEncoded = Left$(sEncoded, Pos) _
             & Mid$(sEncoded, Pos + 3)
   pointer = Pos + 1
Loop
On Error GoTo 0     'reset error handling
UrlDecode = sEncoded
Exit Function

errorUrlDecode:
If Err.Number = 13 Then      'Type Mismatch error
   Err.Clear
   Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
Else
   Err.Raise Err.Number
End If
Resume Next
End Function

Function GetCgiValue(cgiName As String) As String
CGIForm.RequestText = ""

If KIT <> "CMD" Then
Dim n As Integer
For n = 1 To UBound(tPair)
    If UCase$(cgiName) = UCase$(tPair(n).name) Then
       If GetCgiValue = "" Then
          GetCgiValue = tPair(n).Value
       Else             ' allow for multiple selections
          GetCgiValue = GetCgiValue & ";" & tPair(n).Value
       End If
    End If
Next n
End If




If KIT = "CMD" Then
 Dim flL As String
 flL = Split(Command$, "/" & cgiName & "=")(1)
 flL = Split(flL, "/")(0)
 flL = RTrim(flL)
 flL = LTrim(flL)
 GetCgiValue = (flL)
End If

CGIForm.RequestText = GetCgiValue
DoEvents

End Function

Sub SendHeader(sTitle As String)
Send "Status: 200 OK"
Send "Content-type: text/html" & vbCrLf
Send "Cache -Control: Max -age = 31536000" & vbCrLf
'Send "<HTML><HEAD><TITLE>" & sTitle & "</TITLE></HEAD>"
End Sub

Sub SendFooter()
Send "</BODY></HTML>"
End Sub

Sub Send(ByVal s As String)
Dim lBytesWritten As Long

s = s & vbCrLf
WriteFile hStdOut, s, Len(s), lBytesWritten, ByVal 0&

End Sub

Sub SendI(ByVal s As String)
Dim lBytesWritten As Long
WriteFile hStdOut, s, Len(s), lBytesWritten, ByVal 0&

End Sub


Sub SendB(s As String)
Dim lBytesWritten As Long
WriteFile hStdOut, s, Len(s), lBytesWritten, ByVal 0&

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTempPath Lib "kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal _
lpBuffer As String) As Long

Private Const MAX_PATH = 260
Private Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Public Session_Name As String

Function EncodeID(ByVal asContents) As String
        Dim lnPosition
        Dim lsResult
        Dim Char1
        Dim Char2
        Dim Char3
        Dim Char4
        Dim Byte1
        Dim Byte2
        Dim Byte3
        Dim SaveBits1
        Dim SaveBits2
        Dim lsGroupBinary
        Dim lsGroup64
        
        If Len(asContents) Mod 3 > 0 Then asContents = asContents & String(3 - (Len(asContents) Mod 3), " ")
        lsResult = ""
        
        For lnPosition = 1 To Len(asContents) Step 3
            lsGroup64 = ""
            lsGroupBinary = Mid(asContents, lnPosition, 3)
    
            Byte1 = Asc(Mid(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
            Byte2 = Asc(Mid(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
            Byte3 = Asc(Mid(lsGroupBinary, 3, 1))
    
            Char1 = Mid(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
            Char2 = Mid(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
            Char3 = Mid(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
            Char4 = Mid(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
            lsGroup64 = Char1 & Char2 & Char3 & Char4
            
            lsResult = lsResult + lsGroup64
        Next
        
        EncodeID = lsResult
    End Function

Function name(ByVal nameN As String) As String
On Error Resume Next
Dim TempPath As String
nameN = Replace(nameN, "/", "")
nameN = Replace(nameN, "\", "")
nameN = Replace(nameN, ":", "")
nameN = Replace(nameN, ";", "")

If nameN = "" Then
    'get data
    name = Session_Name
    Else
    'set data
    TempPath = GetTmpPath
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
    MkDir TempPath & "w8-system\"
    MkDir TempPath & "w8-system\session\"
    MkDir TempPath & "w8-system\session\" & nameN & "\"
    Session_Name = nameN
    name = nameN
    End If

End Function

Function Id() As String
tempVarIP = Environ("REMOTE_ADDR")
tempVarAgent = Date & Environ("HTTP_USER_AGENT")
tempID = tempVarIP & tempVarAgent

tempID = Replace(tempID, " ", "")
tempID = Replace(tempID, ".", "")
tempID = Replace(tempID, ",", "")
tempID = Replace(tempID, ":", "")
tempID = Replace(tempID, ";", "")
tempID = Replace(tempID, ")", "")
tempID = Replace(tempID, "(", "")
tempID = Replace(tempID, "/", "")
tempID = Replace(tempID, "\", "")
tempID2 = EncodeID(tempID)
Id = tempID2

End Function

Sub Add(ByVal tIp As String, ByVal nData As String)

f1 = FreeFile
Open GetTmpPath & "w8-system\session\" & Session_Name & "\" & Id & ".tmp" For Append As #f1
Print #f1, tIp & "=" & Chr(34) & nData & Chr(34)
Close #f1


End Sub

Function GetData()

Dim a As String
Dim d As String
f1 = FreeFile
dat = GetTmpPath & "w8-system\session\" & Session_Name & "\" & Id & ".tmp"
Open dat For Input As #f1
Do While Not EOF(f1)
 Line Input #f1, a
 a = Replace(a, LCase("<cgib>"), "")
 a = Replace(a, LCase("</cgib>"), "")
 If a <> "" Then d = d & vbCrLf & a
Loop
Close #f1

If IncludeType = 0 Then ScriptVB.ExecuteStatement d
If IncludeType = 1 Then ScriptVB2.ExecuteStatement d

End Function
Function Delete()
On Error Resume Next
Dim dat As String
dat = GetTmpPath & "w8-system\session\" & Session_Name & "\" & Id & ".tmp"
Kill dat
End Function

Function GetTempDir() As String
    TempPath = GetTmpPath
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
    'GetTempDir = GetTmpPath & "w8-system\session\" & Session_Name
    GetTempDir = Session_Name
End Function
Function GetTmpPath() As String

Dim sFolder As String ' Name of the folder
Dim lRet As Long ' Return Value

sFolder = String(MAX_PATH, 0)
lRet = GetTempPath(MAX_PATH, sFolder)

If lRet <> 0 Then
GetTmpPath = Left(sFolder, InStr(sFolder, _
Chr(0)) - 1)
Else
GetTmpPath = vbNullString
End If

End Function





Attribute VB_Name = "cgi4upld"
Public sUploadFileNamess As String
Public bFilePart  As Boolean
Public sBody      As String
Public sUploadDir As String
Public sUploadFilename As String
Public fPair()    As pair

Public sLog
Public Sub AddPairF(sName As String, sValue As String)
Dim n As Long

n = UBound(fPair) + 1
ReDim Preserve fPair(n)
fPair(n).name = sName
fPair(n).Value = sValue

End Sub

Public Function GetFileInfo(sItem As String) As String
Dim x As Long
 
For x = 1 To UBound(fPair)
   If UCase$(sItem) = UCase$(fPair(x).name) Then
      GetFileInfo = fPair(x).Value   'this is what we are looking for
      Exit For
   End If
Next x

End Function

Public Sub ReadLine(sLine As String)
Dim v As Variant
Dim x As Long
Dim Pos As Long
Dim sItem As String

If InStr(sLine, "filename=""") Then   'identifies the line that
   bFilePart = True                   '  contains the file info.
   ReDim fPair(0)                     'initialize fPair()
End If

v = ParseToArray(sLine, "; ")  'split the line by semi-colon into items
      
For x = 0 To UBound(v)         'for each item within the header line
   sItem = (v(x))
   If sItem = "" Then GoTo Iterate              'go to the next item
   Pos = InStr(sItem, "=")
                   
   If bFilePart Then                            'part contains a file
      If InStr(sItem, "filename=""") = 1 Then   'save the original name
         AddPairF Left$(sItem, Pos - 1), _
             Strip(Mid$(sItem, Pos + 1), Chr$(34))
                                                '...and our temp name
         AddPairF "saveAs", _
             TempFile(sUploadDir, "up")
                                                '...and the file size
         AddPairF "fileSize", _
             Trim$(Str$(Len(sBody)))
      
      ElseIf Pos > 0 Then                       'all other pairs
         AddPairF Left$(sItem, Pos - 1), _
             Strip(Mid$(sItem, Pos + 1), Chr$(34))
      
      ElseIf InStr(1, sItem, "Content-type: ", 1) Then
         AddPairF "Content-type", _
            ParseItem(sItem, ": ", 2)
      End If
   
   Else
      If InStr(sItem, "name=") = 1 Then    'pairs containing "name="
         AddPairT Strip(Mid$(sItem, 6), Chr$(34)), _
                  sBody
      
      ElseIf Pos > 0 Then                   'all other pairs
         AddPairT Left$(sItem, Pos - 1), _
             Strip(Mid$(sItem, Pos + 1), Chr$(34))
             
      End If
   End If
Iterate:

If Mid(sItem, 1, 9) = "filename=" Then
  SitemEm = Split(sItem, "filename=")(1)
  SitemEm = Replace(SitemEm, Chr(34), "")
  SitemEm = Replace(SitemEm, "'", "")
  SendE "<script>parent.ccUploadStatus('" & SitemEm & "');</script>"
  End If
Next x
End Sub


Public Sub AddPairT(sName As String, sValue As String)
Dim n As Long

n = UBound(tPair) + 1
ReDim Preserve tPair(n)
tPair(n).name = sName
tPair(n).Value = sValue

End Sub

Public Function MultiPart(sData As String) As String
Dim sLog          As String  'log message
Dim sHeader       As String  'headers
Dim sBoundary     As String  'boundary-string
Dim lBoundary     As Long    'pos of the crlf at end of 1st boundary
Dim lNextBoundary As Long    'start byte of next boundary-string
Dim lLastBoundary As Long    'start byte of last boundary-string
Dim lBody         As Long    'start byte of body
Dim lBodyLen      As Long    'length of body
Dim Pos           As Long    'pos of target used with InStr
Dim osszeg As Long
Dim tom As Long
Dim tom2 As Long

lBoundary = InStr(1, sData, vbCrLf)
sBoundary = Left$(sData, lBoundary - 1)
lLastBoundary = InStr(1, sData, sBoundary & "--")
Pos = lBoundary + 2                             'move past the crlf
egysz = lLastBoundary / 100


ho = 0
osszeg = 0

Do
   ho = ho + 1
    
   
   sData = Mid$(sData, Pos)
   lBody = InStr(1, sData, vbCrLf & vbCrLf) + 4 'identified by 2 crlfs
   If lBody = 4 Then Exit Do                    'should never happen
   
   sHeader = Left$(sData, lBody - 5)
   
   ' find the next boundary string
   ' get the content (sBody) of the data
   lNextBoundary = InStr(lBody, sData, sBoundary)
   lBodyLen = lNextBoundary - lBody - 2        'there is a crlf between
   sBody = Mid$(sData, lBody, lBodyLen)        '    the body & boundary
   
   ReadPart sHeader
   
   Pos = lNextBoundary + Len(sBoundary) + 2
   osszeg = osszeg + Pos
   
   tom = egysz
   tom2 = osszeg / tom
   tom2 = Fix(tom2)
   
   
   
  SendE "<script>parent.ccUploadPercent(" & tom2 & ");</script>"
   
Loop Until lNextBoundary = lLastBoundary

sLog = "Date:         " & Now & "<br>" _
   & "Filename:     " & GetFileInfo("filename") & "<br>" _
   & "Content-Type: " & GetFileInfo("Content-type") & "<br>" _
   & "Size:         " & GetFileInfo("fileSize") & "<br>" _
   & "From:         " & CGI_RemoteAddr & " " & CGI_RemoteHost
   
MultiPart = sLog
SendE "<script>parent.ccUploadPercent(100);</script>"
 sUploadFileNamess = GetFileInfo("filename")
End Function
Public Sub ReadPart(sHeader As String)
Dim mik As Integer
Dim v As Variant
Dim x As Long
 
'split out the individual lines
v = ParseToArray(sHeader, vbCrLf) 'returns an array

bFilePart = False
For x = 0 To UBound(v)            'for each line within the part
   ReadLine (v(x))
   
Next x
 
If bFilePart Then                 'this part contains a file
 duf = FreeFile
 mik = 0
  If sUploadFilename = "" Or sUploadFilename = "auto" Then sUploadFilename = GetFileInfo("filename"): mik = 1
  sUploadFileNamess = sUploadFileNamess & sUploadFilename & Chr(13)
    
  Open sUploadDir & "\" & sUploadFilename For Binary As #duf
   Put #duf, , sBody
   Close #duf
  If mik = 1 Then sUploadFilename = ""
End If
End Sub



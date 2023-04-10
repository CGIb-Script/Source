Attribute VB_Name = "Module1"
Private Hash As New MD5Hash
Private bytBlock() As Byte
    Public cbFile As String


    Public SetMyFile As String
Public iniFileStat As String
Public iniFileLog As String
Public iniFileRefresh As String
Public iniFileCachePages As String
Public iniFileCache As String
Public iniFileMinimize As String
Public KIT As String
Public yDataX
Public CacheServer As String
Public UNIDATAVALUE
Public Cont_Type
Public Declare Function GetTickCount Lib "kernel32" () As Long
Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer
Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer
Function b64Decode(ByVal asContents)
        Dim lsResult
        Dim lnPosition
        Dim lsGroup64, lsGroupBinary
        Dim Char1, Char2, Char3, Char4
        Dim Byte1, Byte2, Byte3
        If Len(asContents) Mod 4 > 0 Then asContents = asContents & String(4 - (Len(asContents) Mod 4), " ")
        lsResult = ""
        For lnPosition = 1 To Len(asContents) Step 4
            lsGroupBinary = ""
            lsGroup64 = Mid(asContents, lnPosition, 4)
            Char1 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 1, 1)) - 1
            Char2 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 2, 1)) - 1
            Char3 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 3, 1)) - 1
            Char4 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 4, 1)) - 1
            Byte1 = Chr(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
            Byte2 = lsGroupBinary & Chr(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
            Byte3 = Chr((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
            lsGroupBinary = Byte1 & Byte2 & Byte3
            lsResult = lsResult + lsGroupBinary
        Next
        b64Decode = lsResult
    End Function
Function b64Encode(ByVal asContents)
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
        b64Encode = lsResult
End Function
Function vb_appPath() As String
On Error Resume Next
Dim tempVar As String
Dim t2 As String
Dim a1 As String
Dim szaM As Integer
tempVar = Environ("PATH_TRANSLATED")
t2 = Environ("PATH_TRANSLATED")
 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
Folytat:
  tempVar = Mid(t2, 1, Len(t2) - szaM)
vb_appPath = tempVar
End Function
Sub CGI_Main()
Dim CacheFile As String
Dim LocalCacheFile As String
Dim CacheFolder As String
Dim iniLog As String
On Error Resume Next
kesment = 0
iniFileMinimize = "off"
iniFileCacheRefresh = 60
iniFileCache = "off"
iniFileLog = "off"
iniFileStat = "off"
'CacheServer = "OFF"
Cont_Type = "text/html"
On Error Resume Next
    Dim tempVar As String


    Dim fileN As Integer
    Dim SC, SCL
    Dim Scripts As Integer
    Dim ContType As Integer
    Dim iniFile As String
    
    KIT = "MEM"
  
    IncludeType = 0
    cbFile = Environ("PATH_TRANSLATED")
    SetMyFile = cbFile
    fileN = FreeFile
    If Command$ <> "" Then
       cbFile = Command$
       cbFile = Split(cbFile, " ")(0)
       KIT = "CMD"
    End If
    ccb = SSearch(LCase(cbFile), ".ccb")
    If ccb = "" Then ccb = 0
    If Fix(ccb) > 0 Then KIT = "CCB"
    If KIT = "CCB" Then
     If Cont_Type = "text/html" Then Send "Status: 200 OK"
     Send "Content-type: " & Cont_Type & vbNewLine
    End If
 If cbFile <> "" And KIT <> "CCB" Then
  
  Rem open ini file %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  iniFile = ""
  tempVarce = ""
  
  On Error Resume Next
  fileN = FreeFile
  
  Open AppPath & "\cc.ini" For Input As #fileN
   For ce = 1 To 40
    Line Input #fileN, tempVarce
    iniFile = iniFile & tempVarce & vbNewLine
   Next
  Close #fileN
  
  tempVarce = ""
  
  CacheDomain = Environ("SERVER_NAME")
  CacheDomain = Replace(CacheDomain, "\", "")
  CacheDomain = Replace(CacheDomain, ".", "")
  CacheDomain = Replace(CacheDomain, ":", "")
  CacheFolder = CacheDomain & appFileName
  CacheFile = Environ("QUERY_STRING")
  
  If CacheFile = "" Then CacheFile = "root"
  LocalCacheFile = Trim(b64Encode(CacheFile))
  CacheFile = Trim(b64Encode(CacheFolder)) & "\" & Trim(b64Encode(CacheFile))

  
  
  
                                     iniFileCache = Split(iniFile, "cache=")(1)
                                     iniFileCache = Split(iniFileCache, ";")(0)
                                     
                                     iniFileCachePages = Split(iniFile, "cache_pages=")(1)
                                     iniFileCachePages = Split(iniFileCachePages, ";")(0)
                                     
                                     iniFileMinimize = Split(iniFile, "minimize=")(1)
                                     iniFileMinimize = Split(iniFileMinimize, ";")(0)
  
                                     iniFileRefresh = Split(iniFile, "cache_refresh=")(1)
                                     iniFileRefresh = Split(iniFileRefresh, ";")(0)
                                     
                                     iniFileLog = Split(iniFile, "access_log=")(1)
                                     iniFileLog = Split(iniFileLog, ";")(0)
                                     
                                     iniFileStat = Split(iniFile, "stat_log=")(1)
                                     iniFileStat = Split(iniFileStat, ";")(0)
                                     
                                     iniFileRefreshPath = Split(iniFile, "cache_refresh_path=")(1)
                                     iniFileRefreshPath = Split(iniFileRefreshPath, ";")(0)
                                     iniFileRefreshPath = Trim(iniFileRefreshPath)
                                     iniFileRefreshPath = Replace(iniFileRefreshPath, Chr(10), "")
                                     iniFileRefreshPath = Replace(iniFileRefreshPath, Chr(13), "")
                                     'If Fix(iniFileRefresh) = 0 Then iniFileRefresh = iniFileCacheRefresh
  
  
  
  If Len(iniFileRefreshPath) < 3 Then
     MkDir AppPath & "\ccini\cache\" & Trim(b64Encode(CacheFolder))
     FullCacheFile = AppPath & "\ccini\cache\" & CacheFile & ".che"
     Else
     MkDir iniFileRefreshPath & "\" & Trim(b64Encode(CacheFolder))
     FullCacheFile = iniFileRefreshPath & "\" & CacheFile & ".che"
  End If
  
  
  checkFile = IsFile(FullCacheFile)
  cachefileDate = FileDateTime(FullCacheFile)
  cachefileDate = DateDiff("n", cachefileDate, Now)
  cachefileDate = Int(cachefileDate)
 
 kesment = 0
 tovabb = "igen"
 
 If iniFileCache = "on" And Int(cachefileDate) >= Int(iniFileRefresh) Then tovabb = "igen": kesment = 1
 If iniFileCache = "on" And checkFile <> "True" Then tovabb = "igen": kesment = 1
 If iniFileCache = "on" And Int(cachefileDate) < Int(iniFileRefresh) Then tovabb = "nem": kesment = 0
 If iniFileCache = "off" Then tovabb = "igen": kesment = 0
  
  
  If iniFileCache = "on" And checkFile = "True" And tovabb = "nem" Then
   c_betolt = 1
   
   If iniFileCachePages <> "" Then
      wok = 0
      c_urlek = SSearch(iniFileCachePages, ",")
      For j = 1 To Fix(c_urlek)
      eke = Split(iniFileCachePages, ",")(j - 1)
       c_keres = SSearch(Environ("QUERY_STRING"), eke)
       If Fix(c_keres) > 0 Then wok = wok + 1
      Next
       If wok = 0 Then c_betolt = 0
   End If

   
   If c_betolt = 1 Then
    fileN = 57
    Open FullCacheFile For Input As #fileN
    Do While Not EOF(fileN)
    Line Input #fileN, tempVar
    vege = vege & tempVar & vbNewLine
    Loop
    Close #fileN
    AllSite = vege
    tovabb = "nem"
   End If
  End If
  
 End If
  
  
 If KIT = "CCB" Then tovabb = "igen"
  
 On Error GoTo 0
  Rem open ini file %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  
  If tovabb = "igen" Then
  AllSite = ""
  fileN = FreeFile
  Open cbFile For Input As #fileN
   Do While Not EOF(fileN)
    Line Input #fileN, tempVar
    If Mid(tempVar, 1, 4) = "[CC]" Then tempVar = b64Decode(Mid(tempVar, 5, Len(tempVar)))
    AllSite = AllSite & tempVar & vbNewLine
   Loop
  Close #fileN
  End If
  
If tovabb = "igen" Then

Load CGIForm
    
OpenFileName = cbFile
Scripts = SSearch(AllSite, "<cgib>")


Rem ### interpreterek
Set ScriptVB = CreateObject("MSScriptControl.ScriptControl")
ScriptVB.Language = "VBScript"
ScriptVB.Reset
ScriptVB.AllowUI = True

 Set ScriptVBRun = CreateObject("MSScriptControl.ScriptControl")
 ScriptVBRun.Language = "VBScript"
 ScriptVBRun.Reset
 ScriptVBRun.AllowUI = True

 Set ScriptVB2 = CreateObject("MSScriptControl.ScriptControl")
 ScriptVB2.Language = "VBScript"
 ScriptVB2.Reset
 ScriptVB2.AllowUI = True

 Set ScriptVB3 = CreateObject("MSScriptControl.ScriptControl")
 ScriptVB3.Language = "VBScript"
 ScriptVB3.Reset
 ScriptVB3.AllowUI = True
Rem ### interpreterek


Rem ### függvények hozzáadása az interpreterekhez
Dim cls As New StandardClass
    ScriptVB.AddObject "Me", cls, True
    ScriptVB2.AddObject "Me", cls, True
    ScriptVB3.AddObject "Me", cls, True
    
Dim clsSession As New Session
    ScriptVB.AddObject "Session", clsSession, False
    ScriptVB2.AddObject "Session", clsSession, False
    ScriptVB3.AddObject "Session", clsSession, False
    
Dim clsApp As New AppData
    ScriptVB.AddObject "App", clsApp, False
    ScriptVB2.AddObject "App", clsApp, False
    ScriptVB3.AddObject "App", clsApp, False
    
Dim clsUTF8 As New UTF8
    ScriptVB.AddObject "UTF8", clsUTF8, False
    ScriptVB2.AddObject "UTF8", clsUTF8, False
    ScriptVB3.AddObject "UTF8", clsUTF8, False
    
Dim clsXLS As New XLS
    ScriptVB.AddObject "XLS", clsXLS, False
    ScriptVB2.AddObject "XLS", clsXLS, False
    ScriptVB3.AddObject "XLS", clsXLS, False
  
  
Rem ### függvények hozzáadása az interpreterekhez

If KIT = "CMD" Then
'REM
End If


 If Fix(Scripts) > 0 Then
  vbsHead = "public Function Import(ByRef obj, ByVal progID) " & vbCrLf
  vbsHead = vbsHead & " Set obj = CreateObject(progID) " & vbCrLf
  vbsHead = vbsHead & " End Function " & vbCrLf
  vbsHead = vbsHead & "public sub Print(Byval str) " & vbCrLf
  vbsHead = vbsHead & "  echo str " & vbCrLf
  vbsHead = vbsHead & " End sub " & vbCrLf

 
  For Interpreter = 1 To Fix(Scripts)
   SC = SCrop(AllSite, "<cgib>", "</cgib>")
   SCL = Replace(SC, "<cgib>", "")
   SCL = Replace(SCL, "</cgib>", "")
   ScriptNow = "": ScriptVB.AddCode SCL & vbCrLf & vbsHead
   AllSite = Replace(AllSite, SC, ScriptNow, 1, Interpreter)
  Next Interpreter
 End If



End If

mehet = 1

If KIT = "MEM" Then
 If Cont_Type = "text/html" Then Send "Status: 200 OK"
 Send "Content-type: " & Cont_Type & vbNewLine
End If
 
 On Error Resume Next
 
 If iniFileMinimize = "on" Then
  AllSite = Replace(AllSite, Chr(13), " ")
  AllSite = Replace(AllSite, Chr(10), " ")
  AllSite = Replace(AllSite, Chr(9), " ")
  For t = 1 To 500
  AllSite = Replace(AllSite, "  ", " ")
  Next
 End If
 

 If kesment = 1 Then
 
 c_betolt = 1
   
   If iniFileCachePages <> "" Then
      wok = 0
      c_urlek = SSearch(iniFileCachePages, ",")
      For j = 1 To Fix(c_urlek)
      eke = Split(iniFileCachePages, ",")(j - 1)
       c_keres = SSearch(Environ("QUERY_STRING"), eke)
       If Fix(c_keres) > 0 Then wok = wok + 1
      Next
       If wok = 0 Then c_betolt = 0
   End If

 
  If c_betolt = 1 Then
  urlm = Environ("QUERY_STRING")
  If iniFileRefreshPath = "" Then MkDir AppPath & "\ccini"
  If iniFileRefreshPath = "" Then MkDir AppPath & "\ccini\cache"
  Close All
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(FullCacheFile)
    Fileout.Write AllSite
    Fileout.Close
    
    Close All
    csumi = 1953
    
  End If
 End If
 

most = Time$
most = Replace(most, ":", "")
most = Mid(most, 1, 4)

 If KIT = "MEM" Then If mehet = 1 Then Send AllSite
 If KIT <> "MEM" Then Close All: End
 
 If Trim(iniFileLog) = "on" And Len(Environ("HTTP_REFERER")) > 3 Then
  f = FreeFile
  Open AppPath & "\access.log" For Append As #f
  Print #f, Now & " -- " & Environ("REMOTE_ADDR") & " -- " & Environ("HTTP_REFERER") & " -- " & Environ("PATH_TRANSLATED")
  Close #f
 End If
 
  If Trim(iniFileStat) = "on" And Len(Environ("HTTP_REFERER")) > 3 Then
   f = FreeFile
   MkDir AppPath & "\stat_log": DoEvents
   MkDir AppPath & "\stat_log\" & Trim(Year(Now)): DoEvents
   MkDir AppPath & "\stat_log\" & Trim(Year(Now)) & "\" & Trim(Month(Now)): DoEvents
   statFileData = AppPath & "\stat_log\" & Trim(Year(Now)) & "\" & Trim(Month(Now)) & "\" & Trim(Day(Now)) & ".bsg"
   Open statFileData For Append As #f
   Print #f, Now & "|" & Environ("REMOTE_ADDR") & "|" & Environ("REQUEST_URI") & "|" & Environ("HTTP_REFERER") & "|" & Environ("HTTP_ACCEPT_LANGUAGE") & "|" & Environ("HTTP_USER_AGENT")
   Close #f
  End If

Close All


 If csumi = 1953 Then
    MkDir AppPath & "\cache_log": DoEvents
    Close All
    f = FreeFile
    Open AppPath & "\cache_log\" & LocalCacheFile & ".bsg" For Output As #f
    Print #f, "cacheFilePath = " & Chr(34) & FullCacheFile & Chr(34)
    Close #f
    Close All
 End If

 
End
End
End
End

End Sub
Public Sub SendE(Data As String)
If KIT = "" Then Send Data
If KIT = "CMD" Then SendI Data
If KIT = "CCB" Then Send Data
If KIT = "MEM" Then ScriptNow = ScriptNow & Data
End Sub
Public Sub SendELine(Data As String)
If KIT = "" Then Send Data
If KIT = "CMD" Then Send Data
If KIT = "CCB" Then Send Data
If KIT = "MEM" Then ScriptNow = ScriptNow & Data
End Sub

Public Function SCrop(ByVal ifi As String, ByVal startStr1 As String, ByVal endStr1 As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As Long
Dim e As Long
Dim h As Long
Dim sS As Long
Dim vV As Long
keresh = Len(startStr1)
s = InStr(1, ifi, startStr1)
e = InStr(s, ifi, endStr1)
sS = s
vV = e - (sS)
tempVar = Mid(ifi, sS, vV + Len(endStr1))
If Err.Description <> "" Then tempVar = ""
SCrop = tempVar
End Function
Public Function SSearch(ByVal ifi As String, ByVal searchStr As String) As Integer
On Error Resume Next
Dim tempVar As String
Dim s As String
Dim jo As Integer
  jo = 0
  For t = 1 To Len(ifi)
   s = Mid(ifi, t, Len(searchStr))
   If s = searchStr Then jo = jo + 1
   Next
 
SSearch = jo
End Function

Public Function CsData(ByVal te As String)
mi = Replace(te, Chr(13), " ")
mi = Replace(mi, Chr(10), "")
mi = Replace(mi, "\n", " ")

For z = 1 To 500
mi = Replace(mi, "  ", " ")
Next

CsData = mi
End Function
Public Function CsEnc(ByVal asContents As String)



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
        
        CsEnc = lsResult
    End Function
Public Function AppPath() As String
On Error Resume Next

If KIT <> "CMD" Then
 Dim tempVar As String
 Dim t2 As String
 Dim a1 As String
 Dim szaM As Integer

 tempVar = Environ("PATH_TRANSLATED")
 t2 = Environ("PATH_TRANSLATED")

 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
End If

 
Folytat:
If KIT <> "CMD" Then
  tempVar = Mid(t2, 1, Len(t2) - szaM)
  AppPath = tempVar
End If



If KIT = "CMD" Then AppPath = App.Path
End Function

Function IsFile(ByVal Path As String) As String
On Error Resume Next
Dim tempVar As String
Open Path For Input As 111
Close 111
If Err.Description = "" Then tempVar = "True"
If Err.Description <> "" Then tempVar = "False"
IsFile = tempVar
End Function

Function appFileName() As String
On Error Resume Next
Dim tempVar As String
Dim t2 As String
Dim a1 As String
Dim szaM As Integer

tempVar = Environ("PATH_TRANSLATED")
t2 = Environ("PATH_TRANSLATED")

 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
 
Folytat:
  tempVar = Mid(t2, Len(t2) - szaM + 2, 100)
appFileName = tempVar

End Function

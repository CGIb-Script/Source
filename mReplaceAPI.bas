Attribute VB_Name = "mReplaceAPI"
Private Type DOCINFOW
        cbSize As Long
        lpszDocName As Long
        lpszOutput As Long
        lpszDatatype As Long
        fwType As Long
End Type

Private Declare Function StartDocW Lib "gdi32" (ByVal hDC As Long, ByRef lpDI As DOCINFOW) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadID As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private mAutoRestoreFunctionOnExit As cAutoRestoreFunctionOnExit

Private mPrinterFilePath As String
Private mVBMainHwnd As Long

Public Property Get PrinterFilePath() As String
    PrinterFilePath = mPrinterFilePath
End Property

Public Property Let PrinterFilePath(nPath As String)
    If nPath <> mPrinterFilePath Then
        mPrinterFilePath = nPath
        If mPrinterFilePath = "" Then
            RestoreStartDocAPIFunction
        Else
            ReplaceStartDocAPIFunction
        End If
    End If
End Property

Public Sub ReplaceFunction(ByVal Dll As String, ByVal Func As String, ByVal Add As Long, ByRef alOld() As Long)
    Dim hmod As Long
    Dim lPtr As Long
    Dim alNew(0 To 2) As Long
    Dim iOldProtect As Long
    
    hmod = LoadLibrary(Dll)
    lPtr = GetProcAddress(hmod, Func)
    
    ' Crate the new ASM intructions block
    alNew(0) = &HB8909090 ' nop/nop/mov eax   (move to the eax register what is in the following address)
    alNew(1) = Add        ' function address  (here goes the addess of the replacement function)
    alNew(2) = &H9090E0FF ' jmp eax/nop/nop   (jump to the addess that is in the eax register)
    
    CopyMemory alOld(0), ByVal lPtr, 12
    VirtualProtect lPtr, 12, PAGE_EXECUTE_READWRITE, iOldProtect
    CopyMemory ByVal lPtr, alNew(0), 12
    VirtualProtect lPtr, 12, iOldProtect, iOldProtect
    FreeLibrary hmod
End Sub

Public Sub RestoreFunction(ByVal Dll As String, ByVal Func As String, ByRef alOld() As Long)
    Dim hmod As Long
    Dim lPtr As Long
    Dim alNew(0 To 2) As Long
    Dim iOldProtect As Long
   
    hmod = LoadLibrary(Dll)
    lPtr = GetProcAddress(hmod, Func)
    VirtualProtect lPtr, 12, PAGE_EXECUTE_READWRITE, iOldProtect
    CopyMemory ByVal lPtr, alOld(0), 12
    VirtualProtect lPtr, 12, iOldProtect, iOldProtect
    FreeLibrary hmod
End Sub

Private Sub ReplaceStartDocAPIFunction()
    If GetProp(GetVBMainHwnd, "STW_API_Replaced") = 0 Then
        Dim iOld(0 To 2) As Long
        Dim c As Long
        
        SetProp GetVBMainHwnd, "STW_API_Replaced", 1
        If mAutoRestoreFunctionOnExit Is Nothing Then Set mAutoRestoreFunctionOnExit = New cAutoRestoreFunctionOnExit
        ReplaceFunction "gdi32.dll", "StartDocA", AddressOf StartDocAReplacementProc, iOld
        For c = 0 To 2
            SetProp GetVBMainHwnd, "STW_API_" & CStr(c), iOld(c)
        Next c
    End If
End Sub

Private Sub RestoreStartDocAPIFunction()
    If GetProp(GetVBMainHwnd, "STW_API_Replaced") = 1 Then
        Dim iOld(0 To 2) As Long
        Dim c As Long
        
        For c = 0 To 2
            iOld(c) = GetProp(GetVBMainHwnd, "STW_API_" & CStr(c))
            RemoveProp GetVBMainHwnd, "STW_API_" & CStr(c)
        Next c
        RestoreFunction "gdi32.dll", "StartDocA", iOld
        RemoveProp GetVBMainHwnd, "STW_API_Replaced"
        Set mAutoRestoreFunctionOnExit = Nothing
    End If
End Sub

Public Sub Terminate()
    RestoreStartDocAPIFunction
End Sub

Private Function StartDocAReplacementProc(ByVal hDC As Long, ByRef lpDI As DOCINFOW) As Long
    lpDI.lpszOutput = StrPtr(mPrinterFilePath)
    StartDocAReplacementProc = StartDocW(hDC, lpDI)
End Function

Private Function GetVBMainHwnd() As Long
    If mVBMainHwnd = 0 Then EnumThreadWindows App.ThreadID, AddressOf EnumThreadProc_GetIDEMainWindow, 0&
    GetVBMainHwnd = mVBMainHwnd
End Function

Private Function EnumThreadProc_GetIDEMainWindow(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim iBuff As String * 255
    Dim iWinClass As String
    Dim iRet As Long
    
    iRet = GetClassName(lhWnd, iBuff, 255)
    
    If iRet > 0 Then
        iWinClass = Left$(iBuff, iRet)
    Else
        iWinClass = ""
    End If
    
    Select Case iWinClass
        Case "ThunderRT6Main"
            mVBMainHwnd = lhWnd
            EnumThreadProc_GetIDEMainWindow = 0
    End Select
    EnumThreadProc_GetIDEMainWindow = 1
End Function




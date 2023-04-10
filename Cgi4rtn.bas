Attribute VB_Name = "CGI4RTN"
Private Declare Function GetTempFileName Lib "kernel32" _
    Alias "GetTempFileNameA" _
   (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long

Public Function Strip(ByVal sSource As String, _
                      ByVal sTarget As String, _
                     Optional vType As Variant) As String
Dim backward As Boolean 'direction of search
Dim x        As Long    'counter
Dim Pos      As Long    'position of sTarget
Dim pointer  As Long    'start of InStr
Dim lTarget  As Long    'length of sTarget
Dim lType    As Long    'vType converted to long
Dim sType    As String  'vType converted to string

If sTarget = "" Then GoTo exitStrip      ' sTarget cannot be empty
lTarget = Len(sTarget)

'validate vType

If IsMissing(vType) Then
   sType = "A"                           ' default = "A" (all)
ElseIf vType = "" Then
   sType = "A"
ElseIf IsNumeric(vType) Then             ' a number was entered
   GoTo numStrip
Else
   sType = Left$(UCase$(vType), 1)       ' Use only the first character
End If

If InStr("ABLT", sType) = 0 Then sType = "A"

Select Case sType
   Case "A" 'all
      Do
         Pos = InStr(1, sSource, sTarget)
         If Pos = 0 Then Exit Do
         sSource = Left$(sSource, Pos - 1) & Mid$(sSource, Pos + lTarget)
      Loop
   Case "B" 'leading and trailing
      Do While InStr(1, sSource, sTarget) = 1
         sSource = Mid$(sSource, lTarget + 1)
      Loop
      sSource = Reverse(sSource)
      sTarget = Reverse(sTarget)
      Do While InStr(1, sSource, sTarget) = 1
         sSource = Mid$(sSource, lTarget + 1)
      Loop
      sSource = Reverse(sSource)
   Case "L" 'leading
      Do While InStr(1, sSource, sTarget) = 1
         sSource = Mid$(sSource, lTarget + 1)
      Loop
   Case "T" 'trailing
      sSource = Reverse(sSource)
      sTarget = Reverse(sTarget)
      Do While InStr(1, sSource, sTarget) = 1
         sSource = Mid$(sSource, lTarget + 1)
      Loop
      sSource = Reverse(sSource)
End Select
GoTo exitStrip                           ' done

numStrip:
lType = CLng(vType)                      ' convert to long
If lType = 0 Then GoTo exitStrip         ' cannot be zero
x = 1
pointer = 1

If lType < 0 Then
   backward = True
   lType = Abs(lType)
   sSource = Reverse(sSource)
   sTarget = Reverse(sTarget)
End If

Do
   Pos = InStr(pointer, sSource, sTarget)
   If Pos = 0 Then Exit Do
   If x = lType Then
      sSource = Left$(sSource, Pos - 1) _
               & Mid$(sSource, Pos + lTarget)
      Exit Do
   End If
   x = x + 1
   pointer = Pos + lTarget
Loop
If backward Then sSource = Reverse(sSource)

exitStrip:
Strip = sSource
End Function

Public Function Translate(ByVal sSource As String, _
                            ByVal sFrom As String, _
                              ByVal sTo As String, _
                         Optional vType As Variant) As String
Dim backward As Boolean 'direction of search
Dim x        As Long    'counter
Dim pointer  As Long    'start of InStr
Dim Pos      As Long    'position of sFrom
Dim lFrom    As Long    'length of sFrom
Dim lTo      As Long    'length of sTo
Dim lType    As Long    'vType converted to long
Dim sType    As String  'vType converted to string

If sSource = "" Or sFrom = "" Then GoTo exitTranslate
lFrom = Len(sFrom)
lTo = Len(sTo)

'validate vType

If IsMissing(vType) Then
   sType = "A"                          'default = "A" (all)
ElseIf vType = "" Then
   sType = "A"
ElseIf IsNumeric(vType) Then
   GoTo numTranslate                    'translate nth occurrence
Else
   sType = Left$(UCase$(vType), 1)      'a string was entered
End If

If InStr("ABLT", sType) = 0 Then sType = "A"

Select Case sType
 Case "A" 'all
   pointer = 1
   Do
     Pos = InStr(pointer, sSource, sFrom)
     If Pos = 0 Then Exit Do
     sSource = Left$(sSource, Pos - 1) & sTo _
              & Mid$(sSource, Pos + lFrom)
     pointer = Pos + lTo
   Loop
 
 Case "B" 'leading and trailing
   pointer = 1
   Do
     Pos = InStr(pointer, sSource, sFrom)
     If Pos <> pointer Then Exit Do
     sSource = Left$(sSource, Pos - 1) & sTo & Mid$(sSource, Pos + lFrom)
     pointer = Pos + lTo
   Loop
   sSource = Reverse(sSource)
   sFrom = Reverse(sFrom)
   sTo = Reverse(sTo)
   pointer = 1
   Do
     Pos = InStr(pointer, sSource, sFrom)
     If Pos <> pointer Then Exit Do
     sSource = Left$(sSource, Pos - 1) & sTo & Mid$(sSource, Pos + lFrom)
     pointer = Pos + lTo
   Loop
   sSource = Reverse(sSource)
   
 Case "L" 'leading
   pointer = 1
   Do
     Pos = InStr(pointer, sSource, sFrom)
     If Pos <> pointer Then Exit Do
     sSource = Left$(sSource, Pos - 1) & sTo & Mid$(sSource, Pos + lFrom)
     pointer = Pos + lTo
   Loop
   
 Case "T" 'trailing
   sSource = Reverse(sSource)
   sFrom = Reverse(sFrom)
   sTo = Reverse(sTo)
   pointer = 1
   Do
     Pos = InStr(pointer, sSource, sFrom)
     If Pos <> pointer Then Exit Do
     sSource = Left$(sSource, Pos - 1) & sTo & Mid$(sSource, Pos + lFrom)
     x = Pos + lTo
   Loop
   sSource = Reverse(sSource)
End Select
GoTo exitTranslate                       'done

numTranslate:
lType = CLng(vType)                      'convert to long
If lType = 0 Then GoTo exitTranslate     'cannot be zero
x = 1
pointer = 1

If lType < 0 Then                        'negative number
  backward = True                        'search from end
  lType = Abs(lType)
  sSource = Reverse(sSource)
  sFrom = Reverse(sFrom)
  sTo = Reverse(sTo)
End If
  
Do
   Pos = InStr(pointer, sSource, sFrom)
   If Pos = 0 Then Exit Do
   If x = lType Then
      sSource = Left$(sSource, Pos - 1) _
         & sTo & Mid$(sSource, Pos + lFrom)
      Exit Do
   End If
   x = x + 1
   pointer = Pos + lFrom
Loop
If backward Then sSource = Reverse(sSource)

exitTranslate:
Translate = sSource
End Function

Public Function InsertA(ByVal sSource As String, _
                              sTarget As String, _
                        Optional vPos As Variant, _
                        Optional vPad As Variant) As String
Dim lSource     As Long   'length of Source
Dim lSize       As Long   'minimum size needed for sSource
Dim Pos         As Long   'vPos converted to long
Dim pad         As String 'vPad converted to string

If sTarget = "" Then GoTo exitInsertA
lSource = Len(sSource)

'validate vPos

If IsMissing(vPos) Then          'default = 0
   Pos = 0
ElseIf Not IsNumeric(vPos) Then
   Pos = 0
ElseIf Abs(vPos) > 500000 Then   'be reasonable
   Pos = 0
Else                             'no negative numbers
   Pos = Abs(CLng(vPos))
End If

'validate vPad

If IsMissing(vPad) Then          'default = " "
   pad = " "
ElseIf vPad = "" Then
   pad = " "
Else
   pad = Left$(CStr(vPad), 1)
End If

lSize = Pos - lSource
If lSize > 0 Then               'pad character will be used
   pad = String(lSize, pad)     'string of pad characters
Else
   pad = ""
End If

sSource = Left$(sSource, Pos) & pad & sTarget _
         & Mid$(sSource, Pos + 1)

exitInsertA:
InsertA = sSource
End Function

Public Function Overlay(ByVal sSource As String, _
                              sTarget As String, _
                        Optional vPos As Variant, _
                        Optional vPad As Variant) As String
Dim lSource     As Long   'length of sSource
Dim lSize       As Long   'minimum size needed for Overlay
Dim lTarget     As Long   'length of sTarget
Dim Pos         As Long   'vPos converted to long
Dim pad         As String 'vPad converted to string

If sTarget = "" Then GoTo exitOverlay 'sTarget cannot be empty
lTarget = Len(sTarget)
lSource = Len(sSource)

'validate pos

If IsMissing(vPos) Then          'default = 1
   Pos = 1
ElseIf Not IsNumeric(vPos) Then
   Pos = 1
ElseIf vPos = 0 Then             'pos cannot be 0
   Pos = 1
ElseIf Abs(vPos) > 1024000 Then  'be reasonable
   Pos = 1
Else                             'no negative numbers
   Pos = Abs(CLng(vPos))
End If

'validate pad

If IsMissing(vPad) Then          'default = " "
   pad = " "
ElseIf vPad = "" Then
   pad = " "
Else
   pad = Left$(CStr(vPad), 1)    'only the first character of pad
End If

lSize = Pos + lTarget - 1        'expand sSource if necessary
If lSize > lSource Then          'pad character will be used
   sSource = sSource & String(lSize - lSource, pad)
End If
Mid$(sSource, Pos, lTarget) = sTarget

exitOverlay:
Overlay = sSource
End Function

Public Function DelStr(sSource As String, _
                        lStart As Long, _
             Optional vLength As Variant) As String
Dim lSource As Long   'length of sSource
Dim lLength As Long   'vLength converted to long

DelStr = sSource
lSource = Len(sSource)
If lStart <= 0 _
Or lStart > lSource _
Or lSource = 0 Then Exit Function
   
If IsMissing(vLength) Then
   DelStr = Left$(sSource, lStart - 1)
   Exit Function
ElseIf Not IsNumeric(vLength) Then
   Exit Function
End If
lLength = CLng(vLength)
If lLength < 1 Then Exit Function
DelStr = Left$(sSource, lStart - 1) _
        & Mid$(sSource, lStart + lLength)
End Function

Public Function ParseCount(sSource As String, sTarget As String) As Long
Dim pointer As Long    'pointer in sSource
Dim Pos     As Long    'position of sTarget
Dim lTarget As Long    'length of sTarget
Dim lSource As Long    'length of sSource

If sSource = "" Then Exit Function     'nothing to count
If sTarget = "" Then sTarget = " "     'sTarget cannot be empty

lTarget = Len(sTarget)
lSource = Len(sSource)
pointer = 1

Do
   Pos = InStr(pointer, sSource, sTarget)
   If Pos = 0 Then Pos = lSource + 1   'last Target
   ParseCount = ParseCount + 1         'increment
   pointer = Pos + lTarget
Loop Until Pos > lSource
End Function

Public Function ParseItem(ByVal sSource As String, _
                          ByVal sTarget As String, _
                                      n As Long) As String
Dim backward As Boolean   'direction of search
Dim pointer  As Long      'pointer in sSource
Dim Pos      As Long      'position of sTarget
Dim x        As Long      'counter
Dim lTarget  As Long      'length of sTarget
Dim lSource  As Long      'length of sSource

If n = 0 Then Exit Function
If sSource = "" Then Exit Function
If sTarget = "" Then sTarget = " "     'sTarget cannot be empty

lTarget = Len(sTarget)
lSource = Len(sSource)
pointer = 1

If n < 0 Then                          'negative value
   backward = True                     'search from end
   n = Abs(n)
   sSource = Reverse(sSource)
   sTarget = Reverse(sTarget)
End If
   
Do
   Pos = InStr(pointer, sSource, sTarget)
   If Pos = 0 Then Pos = lSource + 1   'last item
   x = x + 1                           'increment
   If n = x Then                       'the item being sought
      ParseItem = Mid$(sSource, pointer, Pos - pointer)
      If backward Then ParseItem = Reverse(ParseItem)
      Exit Do                          'done
   End If
   pointer = Pos + lTarget
Loop Until Pos > lSource
End Function

Public Function ParseToArray(sSource As String, _
                             sTarget As String) As Variant

Dim a()      As String  'array containing elements
Dim pointer  As Long    'pointer in sSource
Dim Pos      As Long    'position of sTarget
Dim x        As Long    'array index
Dim lTarget  As Long    'length of sTarget
Dim lSource  As Long    'length of sSource

If sTarget = "" Then sTarget = " "      'sTarget cannot be null

lTarget = Len(sTarget)
lSource = Len(sSource)
pointer = 1

Do
   Pos = InStr(pointer, sSource, sTarget)
   If Pos = 0 Then Pos = lSource + 1            'last item
   ReDim Preserve a(x)                          'add to the array
   a(x) = Mid$(sSource, pointer, Pos - pointer) 'put item in the array
   x = x + 1                                    'increment array index
   pointer = Pos + lTarget                      'skip to the next item
Loop Until Pos > lSource

ParseToArray = a()                  'return the array as a variant
Erase a()
End Function

Public Function Reverse(sSource As String) As String
Dim x       As Long   'counter
Dim lSource As Long   'length of sSource
Dim lPlus   As Long   'lSource + 1

Reverse = sSource
lSource = Len(sSource)
If lSource < 2 Then Exit Function
lPlus = lSource + 1
For x = 1 To lSource
    Mid$(Reverse, lPlus - x, 1) = Mid$(sSource, x, 1)
Next x
End Function

Public Function UrlEncode(ByVal sSource As String) As String
Dim x       As Long   'counter
Dim c       As String 'character
Dim h       As String 'hexadecimal
Dim Pos     As Long   'position used with Instr()
Dim pointer As Long   'pointer in sSource

x = 1
Do Until x > Len(sSource)
   c = Mid$(sSource, x, 1)
   
   If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789.-_* ", c, 1) Then
      x = x + 1
   Else
      'replace reserved chars with "%xx"
      h = Hex$(Asc(c))
      If Len(h) = 1 Then h = "0" & h
      
      sSource = Left$(sSource, x - 1) _
        & "%" & h _
        & Mid$(sSource, x + 1)
      x = x + 3
   End If
Loop

'replace " " with "+"
pointer = 1
Do
   Pos = InStr(pointer, sSource, " ")
   If Pos = 0 Then Exit Do
   Mid$(sSource, Pos, 1) = "+"
   pointer = Pos + 1
Loop

UrlEncode = sSource
End Function

Public Function TempFile(sPath As String, sPrefix As String) As String
Dim x  As Long
Dim rc As Long
End Function


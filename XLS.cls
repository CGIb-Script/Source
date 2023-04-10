VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objDataExcel As Object

Function NewFile()
 On Error Resume Next
 
  Set objDataExcel = CreateObject("Excel.Application")
  objDataExcel.Visible = False
  objDataExcel.DisplayAlerts = False
  objDataExcel.Workbooks.Add
 
 If Err.Description = "" Then
  objExcel.Quit
  Set objExcel = Nothing
  NewFile = "Success!"
  Else
  NewFile = Err.Number & ", " & Err.Description
 End If
  
End Function


Function Add(ByVal cRow As String, ByVal cCol As String, ByVal strText As String)
 On Error Resume Next
 objDataExcel.cells(cRow, cCol).Value = strText
 
 If Err.Description = "" Then
  Add = "Success!"
  Else
  Add = Err.Number & ", " & Err.Description
 End If
 
End Function

Function SaveAs(ByVal fileNeve As String)
 On Error Resume Next
 objDataExcel.ActiveWorkbook.SaveAs fileNeve, FileFormat:=56
 
 If Err.Description = "" Then
   objDataExcel.Quit
   Set objDataExcel = Nothing
   SaveAs = "Success!"
  Else
   SaveAs = Err.Number & ", " & Err.Description
 End If
 
End Function
Function CloseFile()
 On Error Resume Next
   objDataExcel.Quit
   Set objDataExcel = Nothing
 
 If Err.Description = "" Then
   CloseFile = "Success!"
  Else
   CloseFile = Err.Number & ", " & Err.Description
 End If
 
End Function


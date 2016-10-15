Attribute VB_Name = "CheckIfSheetExists"
Option Explicit
Function CheckIfSheetExists(shtName As String, Optional wb As Workbook) As Boolean
    
'Purpose:   Return boolean TRUE if a sheet exists.

Define_Variable:

    Dim sht As Worksheet
    
Check_If_Sheet_Exists:

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
    
End Function

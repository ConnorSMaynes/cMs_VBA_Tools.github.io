Attribute VB_Name = "SecureGSheetQueries"
Option Explicit
Sub RunQueries()

'Purpose:   Use the Google SQL API to drag in data from the google spreadsheet where
'           the data is located and print it on the specified sheet of this workbook
'           in the same columns as it is found in the Google Sheet

Define_Variables:

    Dim URL As String               'store URL of workbook you want to query
    Dim sheetID As String           'store the sheet id of the workbook you want to query
    Dim printToSheet As Worksheet   'worksheet where you are storing your query data
    Dim i As Integer                'iterative variable for columns (query each column at a time -- for large workbooks)
    Dim dataEnd                     'define the last column number in the worksheet
    
Process_Initialize:

    Application.ScreenUpdating = False          'stop screen updating to reduce processing time
    sheetID = " ***put your sheet id here*** "
    URL = "https://docs.google.com/spreadsheets/u/1/d/" & sheetID & "/gviz/tq?tqx=out:html&tq=SELECT+"
    dataEnd = 10
    printToSheet = Sheets("*** your sheet name here ***")
    
Run_Sheet_Query:
    
    i = 1                                       'initialize query to column A of the google sheet
    printToSheet.Cells.ClearContents            'clear previous queries from sheet
    Do While i <= dataEnd
        Call CollectGSheetData(URL, _
                                ConvertToLetter(i), _
                                False, _
                                printToSheet)
        i = i + 1
    Loop
    
End_Statements:

    Application.ScreenUpdating = True
    
End Sub
Function CollectGSheetData(ByVal uAdd As String, _
                ByVal sourceDestCol As String, _
                ByVal fileOpenRefresh As Boolean, _
                ByVal uSht As Worksheet)
                
'Purpose:   Run a web query for a specific column of a google spreadsheet.
'           uAdd = Google SQL API web address
'           sourceDestCol = column from the google sheet to extract
'           fileOpenRefresh = if you want the queries to update every time the workbook opens, set to TRUE
'           uSht = sheet to print query to
                
    With uSht.QueryTables.Add(Connection:= _
        "URL;" & uAdd & sourceDestCol _
        , Destination:=uSht.Range("$" & sourceDestCol & "$1"))
        .Name = "html&tq=SELECT+" & sourceDestCol
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = fileOpenRefresh
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
End Function
Public Function ConvertToLetter(ColumnNumber As Integer) As String
    
'Purpose:   Convert number to letter
    
    If ColumnNumber > 26 Then
        ConvertToLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
        Chr(((ColumnNumber - 1) Mod 26) + 65)
    Else
        ConvertToLetter = Chr(ColumnNumber + 64)
    End If
    
End Function

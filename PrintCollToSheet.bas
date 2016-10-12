Attribute VB_Name = "PrintCollToSheet"
Option Explicit
Public Function PrintCollToSheet(ByVal uColl As Collection, _
                                ByVal uSheet As Worksheet, _
                                ByVal uCol As String, _
                                ByVal uRow As Long, _
                                ByVal commaseparatedYesNo As String, _
                                ByVal indexStart As Long, _
                                ByVal indexEnd As Long, _
                                Optional indexIncrmt As Integer)

'Purpose:   Print a collection either vertically in a column on an excel sheet or comma-separated in the user's cell
'           uCol & uRow are the starting row for printing, commaseparatedYesNo allows for comma-separated printing
'           in the cell specified by uCol & uRow. indexStart & indexEnd allows for index range selection for collection
'           Finally, indexIncrmt allows the user to periodically skip elements

Define_Variables:

    Dim i                                                           'iterative variable
    
Function_Initialize:

    If indexIncrmt = 0 Then                                         'if no index increment value specified, assume a value of 1 for printing all collection items
        indexIncrmt = 1
    End If

Print_Collection:

    If UCase(commaseparatedYesNo) = "YES" Then                      'check if the user would like a comma-separated print of their collection in
        uSheet.Range(uCol & uRow) = uColl.Item(indexStart)          'the starting cell they specified
        i = indexStart + 1                                          'start on element 2 of collection, since the first was already printed
        Do While i <= indexEnd                                      'loop through the specified range of elements in collection
            uSheet.Range(uCol & uRow) = uSheet.Range(uCol & uRow) & ", " & uColl.Item(i)
            i = i + indexIncrmt                                     'make periodic collection element jump
        Loop
    Else                                                            'if comma-separated not selected, print down the column
        i = indexStart
        Do While i <= indexEnd
            uSheet.Range(uCol & uRow) = uColl.Item(i)
            i = i + indexIncrmt                                     'make periodic collection element jump
            uRow = uRow + 1                                         'jump to next cell in next row
        Loop
    End If
    
End Function

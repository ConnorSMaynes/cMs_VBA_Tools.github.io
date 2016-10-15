Attribute VB_Name = "createCollFromSheet"
Option Explicit
Public Function CreateCollFromSheet(ByVal uSheet As Worksheet, _
                                    ByVal startCol As String, _
                                    ByVal startRow As Long, _
                                    ByVal uniquesYesNo As String, _
                                    Optional ignoreBlanksYesNo As String) As Collection
      
'Purpose:   Return, as a collection, the values of cells in a user-defined sheet and workbook.
'           The default is to ignore blanks. Calls on function to check if string already exists
'           in the collection when generating a uniques collection
      
Define_Variables:

    Dim endRow As Long          'the last filled row in the user's sheet & column
    Dim startCell As String     'the Cell to start on in the sheet and column, generated from the user inputs
    Dim activeCell As Range     'the cell we are considering adding to collection
    Dim uColl As New Collection 'intermediary collection until we finish creating the collection

Create_Collection:
    
    startCell = startCol & startRow
    endRow = uSheet.Range(startCol & 1048576).End(xlUp).Row                         'find the last filled row number in startCol
    If UCase(ignoreBlanksYesNo) = "YES" Or ignoreBlanksYesNo = vbNullString Then    'default is to ignore blanks
        If UCase(uniquesYesNo) = "NO" Then                                          'if the user wants a collection and not a uniques collection
            For Each activeCell In uSheet.Range(startCell & ":" & startCol & endRow)
                If activeCell.Text <> vbNullString Then                             'check to make sure we are not adding an empty value
                    uColl.Add UCase(activeCell.Text)                                'return collection of all entries, ignoring blank entries
                End If                                                              'between startRow and endRow
            Next activeCell
        ElseIf UCase(uniquesYesNo) = "YES" Then                                     'if the user wants a collection of unique items from the sheet
            For Each activeCell In uSheet.Range(startCell & ":" & startCol & endRow)
                If activeCell.Text <> vbNullString Then                             'check to make sure we are not adding an empty value
                    If Qfuncs.CheckForStringInColl(uColl, activeCell.Text) = 0 Then      'check to make sure the item does not already exist in our collection
                        uColl.Add UCase(activeCell.Text)                            'generate collection of unique entries from data in column
                    End If
                End If
            Next activeCell
        End If
    ElseIf UCase(ignoreBlanksYesNo) = "NO" Then
        If UCase(uniquesYesNo) = "NO" Then                                          'if the user wants a collection and not a uniques collection
            For Each activeCell In uSheet.Range(startCell & ":" & startCol & endRow)
                uColl.Add UCase(activeCell.Text)                                    'return collection of all entries, including blanks
            Next activeCell
        ElseIf UCase(uniquesYesNo) = "YES" Then                                     'if the user wants a collection of unique items from the sheet
            For Each activeCell In uSheet.Range(startCell & ":" & startCol & endRow)
                If Qfuncs.CheckForStringInColl(uColl, activeCell.Text) = 0 Then          'check to make sure the item does not already exist in our collection
                    uColl.Add UCase(activeCell.Text)                                'we must use return a standard format, uppercase. add to collection
                End If
            Next activeCell
        End If
    End If
    Set CreateCollFromSheet = uColl
    
End Function
Public Function CheckForStringInColl(ByVal mySet As Collection, ByVal myCheck As String) As Long
    
'Purpose:   Check if the inputted string exists in any of the collection items
'           return 'TRUE' if it does or 'FALSE' if it does not. NOT CASE SENSITIVE.
'           Checks whether some part of string items in the collection contains myCheck.
'           You CAN CHECK FOR PARTIALS. Return collection item number if myCheck is found.
    
Define_Varaibles:
    
    Dim elm                                                 'iteration variable for elements of collection
    Dim position As Long
    
Function_Initialize:
    
    CheckForStringInColl = 0                                'return zero if the string is not found
    
Check_Collection_For_Users_String:

    position = 1
    For Each elm In mySet
        If InStr(UCase(elm), UCase(myCheck)) > 0 Then       'if the user's string exists in an item, return true. Capitalize to avoid missing
            CheckForStringInColl = position                 'return the position if the string is found
            Exit Function
        End If
        position = position + 1
    Next
    
End Function

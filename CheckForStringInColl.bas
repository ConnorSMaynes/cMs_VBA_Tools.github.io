Attribute VB_Name = "CheckForStringInColl"
Option Explicit
Public Function CheckForStringInColl(ByVal mySet As Collection, _
                                        ByVal myCheck As String) As Long
    
'Purpose:   Check if the inputted string exists in any of the collection items
'           return 'TRUE' if it does or 'FALSE' if it does not. NOT CASE SENSITIVE.
'           Checks whether some part of string items in the collection contains myCheck.
'           You CAN CHECK FOR PARTIALS. Return collection item number if myCheck is found.
    
Define_Varaibles:
    
    Dim elm                                                 'iteration variable for elements of collection
    Dim position As Long
    
Function_Initialize:
    
    CheckForStringInColl = 0                                'return zero if the string is not found.
    
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

Attribute VB_Name = "CreateCollFromListBox"
Option Explicit
Function CreateCollFromListBox(uListBox As Object) As Collection

'Purpose:   Return collection of user selections from specified listbox. Return error message
'           if no listbox selection has been made.

Define_Variables:

    Dim i                                       'iterative variable for looping through listbox
    Dim SelColl As New Collection               'temporary collection for storage of the selected items in the listbox
      
Create_And_Return_Collection:

    Do While i < uListBox.ListCount
        On Error GoTo stop_function             'if the user has not made a selection from the specified listbox, cancel process
        If uListBox.Selected(i) = True Then     'if the item is selected, then add it to the collection
            SelColl.Add uListBox.List(i)
        End If
        i = i + 1
    Loop
    Set CreateCollFromListBox = SelColl
    
stop_function:
    
    If SelColl.Count = 0 Then                   'if there were no selections return error to user and end all processes
        MsgBox ("A selection from a listbox is missing.")
        End
    End If
    
End Function

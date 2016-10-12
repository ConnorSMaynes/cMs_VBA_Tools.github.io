Attribute VB_Name = "FingerSlipCheck"
Option Explicit
Public Function FingerSlipCheck(ByVal actionString As String)

'Purpose:   Generate an question box to ask the user if they would like to continue with the process
'           they selected. Good if the process takes an especially long period of time.

Check_For_User_Error:
    
    If MsgBox("Are you sure you want to " & actionString & " ?", vbYesNo + vbQuestion) = vbNo Then
        End
    End If

End Function

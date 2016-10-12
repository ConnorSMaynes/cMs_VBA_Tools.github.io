Attribute VB_Name = "CvtNumToLet"
Option Explicit
Function CvtNumToLet(ByVal iCol As Integer) As String
   
'Purpose:   convert numbers to letters
   
Define_Variables:

   Dim iAlpha As Integer                                    'store integer value from letter
   Dim iRemainder As Integer                                'store remainder value from letter division
   
Function_Initialize:
   
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   
Convert_Number_To_Letter:

   If iAlpha > 0 Then
      CvtNumToLet = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      CvtNumToLet = CvtNumToLet & Chr(iRemainder + 64)
   End If
   
End Function

Attribute VB_Name = "CvtCollToArr"
Option Explicit
Function CvtCollToArr(uColl As Collection) As Variant()

'Purpose:   Convert a collection to an array and return array.

Define_Variables:

    Dim uArr() As Variant                   'intermediate array
    Dim i As Integer                        'iterative array and collection variable
    
Convert_Collection_to_Array:

    For i = 1 To uColl.Count
        uArr(i - 1) = uColl.Item(i)
    Next
    CvtCollToArr = uArr                     'return array
    
End Function

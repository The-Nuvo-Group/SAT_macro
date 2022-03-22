Attribute VB_Name = "Temp_TestHeader"
'Find the largest word in an array of strings
'Input: Array, an array of strings
'Output: Integer, index of the largest word
Function indexLargestWord(ByRef arrayofwords As Variant) As Integer
    Dim IDX As Integer
    Dim tmpWord As String
    
    'Find Largest Word and it index
    tmpWord = arrayofwords(0)
    For i = LBound(arrayofwords) To UBound(arrayofwords)
        If Len(arrayofwords(i)) > Len(tmpWord) Then
            tmpWord = arrayofwords(i)
            IDX = CInt(i)
        End If
    Next i
    indexLargestWord = IDX
End Function

'Arrange headers in a "saquare" like structure
'Input: String, a header in string format
'Output: String, same header but new-lines and spaces
Function arrangeHeader(header As String) As String

    Dim finalHeader, tmpWord As String
    Dim largest, lrgIndx, i As Integer
    Dim arrayofword() As String
    
    'Split string by spaces
    arrayofwords = Split(header)
    
    lrgIndx = indexLargestWord(arrayofwords)
    
    'Re-arrange words to fit size
    tmpWord = ""
    For i = LBound(arrayofwords) To UBound(arrayofwords)
        If (Len(tmpWord) + Len(arrayofwords(i))) <= Len(arrayofwords(lrgIndx)) Then
            If Len(tmpWord) = 0 Then
                tmpWord = arrayofwords(i)
            Else
                tmpWord = tmpWord + " " + arrayofwords(i)
            End If
        Else
            finalHeader = finalHeader + tmpWord + Chr(10)
            tmpWord = arrayofwords(i)
            'Check if we are at the end of array
            If (i = UBound(arrayofwords)) Then
                finalHeader = finalHeader + arrayofwords(i)
            End If
        End If
    Next i
    
    arrangeHeader = finalHeader

End Function

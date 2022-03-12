Attribute VB_Name = "Temp_TestHeader"
Function arrangeHeader(header As String) As String

    Dim finalHeader, tmp As String
    Dim largest, lrgIndx, i As Integer
    
    
    arrayofWords = Split(header)
    
    'Find Largest Word
    tmp = arrayofWords(0)
    For i = LBound(arrayofWords) To UBound(arrayofWords)
        If Len(arrayofWords(i)) > Len(tmp) Then
            tmp = arrayofWords(i)
            lrgIndx = i
        End If
    Next i
    
    Debug.Print lrgIndx; " "; arrayofWords(lrgIndx)
    Debug.Print ("**")
    
    tmp = ""
    For i = LBound(arrayofWords) To UBound(arrayofWords)
        If Len(arrayofWords(i)) < Len(arrayofWords(lrgIndx)) Then
            If (Len(tmp) + Len(arrayofWords(i))) < Len(arrayofWords(lrgIndx)) Then
                tmp = tmp & " " & arrayofWords(i)
            Else
                finalHeader = finalHeader & " " & tmp & Chr(10)
                tmp = arrayofWords(i)
            End If
        Else
            finalHeader = finalHeader & tmp & Chr(10)
            If (i + 1) < UBound(arrayofWords) Then
                tmp = arrayofWords(i + 1)
            Else
                tmp = ""
            End If
        End If
    Next i
    
    arrangeHeader = finalHeader

End Function


        
    

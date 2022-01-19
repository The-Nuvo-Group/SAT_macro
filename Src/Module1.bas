Attribute VB_Name = "Module1"
Sub Main()
    Dim Annex As New annexAone
    Dim pck As New Collection
    Dim pcv As New Collection
    
    pck.Add "or"
    pck.Add "sc"
    pck.Add "size"
    pck.Add "pq"
    
    pcv.Add 1
    pcv.Add 40
    pcv.Add 1
    pcv.Add 600
    
    'Annex.test = 45
    Debug.Print "Count:" + Str(Annex.getCount)
    Annex.assembleAnnexA1 "PageConfig", pck, pcv
    Debug.Print "Count:" + Str(Annex.getCount)
    Annex.printHashT
    
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "annexAone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Holds Annex A-1 Configuration
Public annexA As Object

'Get Page Confiugration
Public Sub assembleAnnexA1(nameHasht As String, keys As Collection, values As Collection)

    Dim configItems As Object
    Set configItems = CreateObject("Scripting.Dictionary")
    
    If Not (keys.Count = values.Count) Then
        Debug.Print "Key and Values do not match in length"
        
        'Warning and Exit
        Call MsgBox("In assembleAnnexA1Config; Keys and Values diff lenght.", vbQuestion + vbMsgBoxRight, "Parameter Error")
        Exit Sub
    End If
    
    For idx = 1 To keys.Count
        If Not configItems.Exists(keys(idx)) Then
            configItems.Add keys(idx), values(idx)
        End If
    Next
        
    'Add to AnnexA1
    If Not annexA.Exists(nameHasht) Then
        annexA.Add nameHasht, configItems
    End If
    
    'Clear/Free Memory
    Set configItems = Nothing

End Sub

Public Property Get getCount() As Integer
    
    getCount = annexA.Count
    
End Property

Public Sub printHashT()

    For Each k In annexA.keys
        Debug.Print k
        For Each subK In annexA(k).keys
            Debug.Print subK, annexA(k)(subK)
        Next
    Next
    
End Sub


Private Sub Class_Initialize()
    Set annexA = CreateObject("Scripting.Dictionary")
    
End Sub


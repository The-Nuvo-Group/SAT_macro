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
Public annexA As New Dictionary

'Get Page Confiugration
Public Sub assembleAnnexA1(configPath As String)

    Dim FSO As New FileSystemObject
    Dim textstrm As TextStream
    Dim config As Dictionary
    
    Set textstrm = FSO.OpenTextFile(configPath, ForReading)
    
    Set annexA = JsonConverter.ParseJson(textstrm.ReadAll)
    
    textstrm.Close

End Sub

Public Property Get getCount() As Integer
    
    getCount = annexA.Count
    
End Property

Public Sub printHashT()

    For Each k In annexA.keys
        Debug.Print k
        For Each subK In annexA(k).keys
            Debug.Print "  ", subK, "  ", annexA(k)(subK)
        Next
    Next
    
End Sub


Private Sub Class_Initialize()
    'Set annexA = CreateObject("Scripting.Dictionary")
    
End Sub


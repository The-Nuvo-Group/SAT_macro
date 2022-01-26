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
Private annexA As New Dictionary

'Get Page Configuration
Public Sub readConfig(configPath As String)

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

Public Sub setupAnnexPages(ws As Worksheet)

    Debug.Print "WORKING MODE START"
    Call sectionPageSetup(ws, annexA("pagesetup"))
    Call sectionMarginsSetup(ws, annexA("marginsetup"))
    Call sectionHeaderFooterSetup(ws, annexA("headerfootersetup"))
    Debug.Print "WORKING MODE END"
    

End Sub

Private Sub sectionPageSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP PAGE"
    
    ' -- Orientation -- '
    ws.pageSetup.Orientation = config("orientation")
    
    ' -- Paper Size -- '
    ws.pageSetup.PaperSize = config("size")
    
    ' -- Scaling -- '
    ws.pageSetup.Zoom = config("scaling")
    
    ' -- Print Quiality -- '
    ws.pageSetup.PrintQuality = config("quality")
    
End Sub

Private Sub sectionMarginsSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP MARING"
    
    ' -- Set Margins -- '
    ws.pageSetup.TopMargin = Application.InchesToPoints(config("top"))
    ws.pageSetup.BottomMargin = Application.InchesToPoints(config("bottom"))
    
    ws.pageSetup.LeftMargin = Application.InchesToPoints(config("left"))
    ws.pageSetup.RightMargin = Application.InchesToPoints(config("right"))
    
    ws.pageSetup.HeaderMargin = Application.InchesToPoints(config("top"))
    ws.pageSetup.FooterMargin = Application.InchesToPoints(config("top"))
    
End Sub

Private Sub sectionHeaderFooterSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP HEADER/FOOTER"
    ' -- Set Margins -- '
    ws.pageSetup.HeaderMargin = Application.InchesToPoints(config("header"))
    ws.pageSetup.FooterMargin = Application.InchesToPoints(config("footer"))
    
    ' -- Diff odd/even pages -- '
    ws.pageSetup.OddAndEvenPagesHeaderFooter = config("settings")("OddEvenPage")

    ' -- Diff 1st page -- '
    ws.pageSetup.DifferentFirstPageHeaderFooter = config("settings")("DiffFirstPG")

    ' -- Scale with Doc -- '
    ws.pageSetup.ScaleWithDocHeaderFooter = config("settings")("scale")

    ' -- Margin Aligment -- '
    ws.pageSetup.AlignMarginsHeaderFooter = config("settings")("aling")

End Sub




Public Sub printConfig()
    printHashT annexA

End Sub

'Traverses a HashTable in a recursive way
Private Sub printHashT(dict As Dictionary, Optional ByVal level = 0)
    Dim indent As String
    indent = Space(level)
    
    For Each k In dict.keys
        If VarType(dict(k)) = 9 Then ' vbObject == 9
            Debug.Print indent; k; ":  "
            level = level + 1
            printHashT dict(k), level
            level = level - 1
        Else
            Debug.Print indent; "-"; k; ": "; dict(k)
        End If
    Next
        
End Sub



Private Sub Class_Initialize()
    'Set annexA = CreateObject("Scripting.Dictionary")
    
End Sub


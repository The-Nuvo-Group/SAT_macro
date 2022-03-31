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
    Call sectionRepeatRowColSetup(ws)
    Debug.Print "WORKING MODE END"
    

End Sub

Private Sub sectionPageSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP PAGE"
    
    ' -- Orientation -- '
    ws.PageSetup.Orientation = config("orientation")
    
    ' -- Paper Size -- '
    ws.PageSetup.PaperSize = config("size")
    
    ' -- Scaling -- '
    ws.PageSetup.Zoom = config("scaling")
    
    ' -- Print Quiality -- '
    ws.PageSetup.PrintQuality = config("quality")
    
End Sub

Private Sub sectionMarginsSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP MARGIN"
    
    ' -- Set Margins -- '
    ws.PageSetup.TopMargin = Application.InchesToPoints(config("top"))
    ws.PageSetup.BottomMargin = Application.InchesToPoints(config("bottom"))
    
    ws.PageSetup.LeftMargin = Application.InchesToPoints(config("left"))
    ws.PageSetup.RightMargin = Application.InchesToPoints(config("right"))
    
    ws.PageSetup.HeaderMargin = Application.InchesToPoints(config("top"))
    ws.PageSetup.FooterMargin = Application.InchesToPoints(config("top"))
    
End Sub

Private Sub sectionHeaderFooterSetup(ws As Worksheet, config As Dictionary)
    
    Debug.Print "  SETTING UP HEADER/FOOTER"
    ' -- Set Margins -- '
    ws.PageSetup.HeaderMargin = Application.InchesToPoints(config("header"))
    ws.PageSetup.FooterMargin = Application.InchesToPoints(config("footer"))
    
    ' -- Diff odd/even pages -- '
    ws.PageSetup.OddAndEvenPagesHeaderFooter = config("settings")("OddEvenPage")

    ' -- Diff 1st page -- '
    ws.PageSetup.DifferentFirstPageHeaderFooter = config("settings")("DiffFirstPG")

    ' -- Scale with Doc -- '
    ws.PageSetup.ScaleWithDocHeaderFooter = config("settings")("scale")

    ' -- Margin Aligment -- '
    ws.PageSetup.AlignMarginsHeaderFooter = config("settings")("aling")

End Sub

Private Sub sectionRepeatRowColSetup(ws)
    
    Debug.Print "  SETTING UP ROW/COLUMN REPEAT"
    
    'Check if Panes are frozen
    If ActiveWindow.FreezePanes = False Then
        Dim results As String
        results = FreezePanelTarget("A1", "H7", "Property Name")
        
        'Freez  Selected Column'
        Range(results).Select
        
        'Activate Freez Panes if ANY'
        ActiveWindow.FreezePanes = True
    End If
        
    'Split Intersectioin into Row and Colum'
    Dim dict As Object
    Set dict = splitAddress(Cells(ActiveWindow.SplitRow, ActiveWindow.SplitColumn).address)
    
    'Store info about repeating rows and cols into annexA
    annexA.Add "repeating", dict
    
    'Set Repeating Rows and Columns'
    ws.PageSetup.PrintTitleRows = GenerateColRowTitleRange(dict.item("Row"), "Row")
    ws.PageSetup.PrintTitleColumns = GenerateColRowTitleRange(dict.item("Col"), "Col")

End Sub

Public Sub setupPrintArea(ws As Worksheet)

    Debug.Print "  SETTING UP PRINT AREA"
    
    Dim RepCols_width As Double
    Dim dimensions As New Dictionary
    
    'AutoFit all columns and rows in Annex A
    Columns(GenerateColRowTitleRange(lastCl(ws), "Col")).AutoFit
    Rows(GenerateColRowTitleRange(CStr(lastRw(ws)), "Row")).AutoFit
    
    'Set up Reparing columns
    'setupPrintArea_RepCol ws, "A", annexA("repeating").item("Col")
    'setupPrintArea_RepCol ws, "A", lastCl(ws)
    
    'Get width of RepCols
    RepCols_width = CDbl(Format(Range(GenerateColRowTitleRange(annexA("repeating").item("Col"), "Col")).Width / Application.InchesToPoints(1), "0.00"))
    'scale & store down width of RepCols
    dimensions.Add "repCols", scaleDown(RepCols_width, CDbl(annexA("pagesetup")("scaling")))
    
    Debug.Print "REP COLS DIMENSION: "; dimensions("repCols")
    Debug.Print "REP COLS + MRGS DIMENSION: "; Format((annexA("marginsetup")("left") * 2 + dimensions("repCols")), "0.00")
    Debug.Print "Dynamic Cols DIMENSION: "; 11 - (annexA("marginsetup")("left") * 2 + dimensions("repCols"))
    'Debug.Print Application.InchesToPoints(2.5)
    'dimensions.Add "dynCols", Format(((annexA("marginsetup")("left") * 2) + dimensions("repCols") - 11), "0.00"
    printHashT dimensions
    Debug.Print "@@@@"
    'Set up rest of print area
    
    

End Sub

Private Sub setupPrintArea_RepCol(ws As Worksheet, start As String, limit As String)
    
    Dim largestCell, cellRange, headerRange As String
    Dim lenHeader, lenCell As Integer
    
    lenCell = 0
    
    'Loop through columns
    For i = Asc(start) To Asc(limit)
        headerRange = Chr(i) & "1"
        lenHeader = Len(Range(headerRange).text)
        'Create the cell Range for each column starting from cell 2(taking out header cell)
        cellRange = Chr(i) & "2:" & Chr(i) & CStr(lastRw(ws))
        For Each c In Range(cellRange).Cells
            If (Len(c.text) > lenCell) Then
                lenCell = Len(c.text)
                largestCell = c.address
            End If
        Next
        
        'Compare header and cell length
        If (lenHeader > lenCell) Then
            'Arrange header width
            Range(headerRange).Value = arrangeHeader(Range(headerRange).Value)
        End If
        'Autofit current column
        Columns(Chr(i)).AutoFit
    Next
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


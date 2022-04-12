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
    setupPrintArea_RepCol ws, "A", lastCl(ws)
    
    ' --- CALCULATIONS ---
    'Get width of RepCols, scale down & store  width of:
    '   Repeating Columns,
    '   Repeating Columns with left & right annexes borders, and
    '   Dynamic Columns
    RepCols_width = CDbl(Format(Range(GenerateColRowTitleRange(annexA("repeating").item("Col"), "Col")).Width / Application.InchesToPoints(1), "0.00"))
    dimensions.Add "repCols", scaleDown(RepCols_width, CDbl(annexA("pagesetup")("scaling")))
    dimensions.Add "repColsMrgs", Format((annexA("marginsetup")("left") * 2 + dimensions("repCols")), "0.00")
    dimensions.Add "dynCols", 11 - dimensions("repColsMrgs")
    ' --- CALCULATIONS ---
    
    'Set up Dynamic columns
    'setupPrintArea_DynCol ws, annexA("repeating")("Col"), dimensions("dynCols")
    
End Sub

Private Sub setupPrintArea_DynCol(ws As Worksheet, ByVal columnSetoff As String, y As Double)
    Debug.Print "####"
'    Dim tmp As String
'    Debug.Print "Setoff Column: "; Chr(columnSetoff); columnSetoff
'    Debug.Print "Last Column in Annex: "; lastCl(ws); Asc(lastCl(ws))
'    Debug.Print y
'    Debug.Print Asc("Z")
'    tmp = Chr(columnSetoff) + ":" + lastCl(ws)
'    tmp = lastCl(ws) + ":" + lastCl(ws)
'    Debug.Print Range(tmp).Column
    
    Dim start, lmt As Integer
    
    start = Range((columnSetoff + ":" + columnSetoff)).Column + 1
    lmt = Range((lastCl(ws) + ":" + lastCl(ws))).Column
    
    Dim cw, tcw, delta, spreadDelta  As Double
    Dim numCols, colCounter, endSubset, strSubset As Integer
    'Dim strSubset As Integer
    tcw = 0#
    colCounter = start
    strSubset = start
    
    Debug.Print "START: "; start
    Debug.Print "END: "; lmt

    For i = start To lmt
'        cw = Format(Columns(i).Width / Application.InchesToPoints(1), "0.00")
'        Debug.Print cw
        cw = Format(Columns(i).ColumnWidth / Application.InchesToPoints(1), "0.00")
        tcw = tcw + cw
        Debug.Print Columns(i).address; " - Widht: "; cw; " - Total wc: "; tcw; " - Length: "; y
        If tcw > y Then
            tcw = Format(tcw - cw, "0.00")
            delta = Format(y - tcw, "0.00")
            spreadDelta = Format((i - colCounter) / delta, "0.00")
            Debug.Print "i: "; i; " - colCounter: "; colCounter; " (i-colCounter): "; (i - colCounter)
            Debug.Print "**tcw: "; tcw
            Debug.Print "**Delta: "; delta
            Debug.Print "**spreadDelta: "; spreadDelta
            adjustWidths strSubset, i - 1, spreadDelta
            strSubset = i
            twc = 0
        End If
    Next i
    
    Debug.Print "####"
    
End Sub

Private Sub adjustWidths(colStart As Integer, colEnd As Integer, amount As Double)
    Debug.Print "ADJUSTING WIDTHS"
    For i = colStart To colEnd
        Columns(i).ColumnWidth = Format(Columns(i).ColumnWidth / Application.InchesToPoints(1), "0.00") + amount
    Next i
End Sub


Private Sub setupPrintArea_RepCol(ws As Worksheet, startCol As String, limitCol As String)

    Dim largestCell, cellRange, headerRange As String
    Dim start, limit, lenHeader, lenCell As Integer
    
    start = Range((startCol + ":" + startCol)).Column
    limit = Range((limitCol + ":" + limitCol)).Column

    lenCell = 0
        
    'Loop through columns
    For i = start To limit
        headerRange = Split(Cells(1, i).address, "$")(1) & "1"
        lenHeader = Len(Range(headerRange).text)
        'Create the cell Range for each column starting from cell 2(taking out header cell)
        cellRange = Split(Cells(1, i).address, "$")(1) & "2:" & Split(Cells(1, i).address, "$")(1) & CStr(lastRw(ws))
        For Each C In Range(cellRange).Cells
            If (Len(C.text) > lenCell) Then
                lenCell = Len(C.text)
                largestCell = C.address
            End If
        Next

        'Compare header and cell length
        If (lenHeader > lenCell) Then
            'Arrange header width
            Range(headerRange).Value = arrangeHeader(Range(headerRange).Value)
        End If
        'Autofit current column
        Columns(Split(Cells(1, i).address, "$")(1)).AutoFit
        'Reset lenCell for next Column
        lenCell = 0
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_annexAone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Class config objects,variables
Private annexA As New Dictionary
Private anxMessages As New C_Messages
Private arrangeHeadersFlag As Integer


'************************************************************
'Purpose:       Constructor of the class
'Input:         Nothing
'               Example: Class_Initialize()
'Output:        Nothing
'************************************************************
Private Sub Class_Initialize()
    'Set annexA = CreateObject("Scripting.Dictionary")
End Sub


'************************************************************
'Purpose:       Activates or deactivates the header flag.
'Input:         Integer, 1 or 0 indicating status of the flag.
'               Example: setArrangeHeadersFlag()
'Output:        Nothing
'************************************************************
Public Function setArrangeHeadersFlag(flg As Integer)
    arrangeHeadersFlag = flg
    
    anxMessages.set_message ("Arrange Headers Flag set to 1")
End Function


'************************************************************
'Purpose:       Gets the status of the header flag.
'Input:         Nothing
'               Example: getArrangeHeadersFlag()
'Output:        Integer, a 1 or 0
'               Example: 1
'************************************************************
Public Function getArrangeHeadersFlag()
    getArrangeHeadersFlag = arrangeHeadersFlag
End Function


'************************************************************
'Purpose:       Reads and stores JSON configuration file holding profiles.
'Input:         String, path to a JSON file
'               Example: readConfig("C://PATH/TO/A/FILE")
'Output:        Nothing
'************************************************************
'Get Page Configuration
Public Sub readConfig(configPath As String)

    Dim FSO As New FileSystemObject
    Dim textstrm As TextStream
    Dim config As Dictionary
    
    Set textstrm = FSO.OpenTextFile(configPath, ForReading)
    
    Set annexA = JsonConverter.ParseJson(textstrm.ReadAll)
    
    textstrm.Close

End Sub


'************************************************************
'Purpose:       Retrives the value of key count in the Dict annexA.
'Input:         Nothing
'               Example: getCount()
'Output:        Integer
'               Example: 2
'************************************************************
Public Property Get getCount() As Integer
    
    getCount = annexA.Count
    
End Property



'************************************************************
'Purpose:       Calls subrutines that modify an excel sheet.
'Input:         Worksheet, active worksheet to be modify
'               Example: setupAnnexPages(activeWorksheet)
'Output:        Nothing
'************************************************************
Public Sub setupAnnexPages(ws As Worksheet)

    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("WORKING MODE START")
    
    Call sectionPageSetup(ws, annexA("pagesetup"))
    Call sectionMarginsSetup(ws, annexA("marginsetup"))
    Call sectionHeaderFooterSetup(ws, annexA("headerfootersetup"))
    Call sectionRepeatRowColSetup(ws)
    Call SectionPrintArea(ws)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("WORKING MODE END")
    
    tmp = MsgBox("All done, you can close the Form now.", vbInformation)
End Sub


'************************************************************
'Purpose:       Subrutine that sets an excel sheet's Page Setup Properties
'Input:         Worksheet, active worksheet to be modify
'               Dictionary, dict with parameters to be set
'               Example: sectionPageSetup(activeWorksheet, DictProperties)
'Output:        Nothing
'************************************************************
Private Sub sectionPageSetup(ws As Worksheet, config As Dictionary)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("  SETTING UP PAGE")
    
    ' -- Orientation -- '
    ws.PageSetup.Orientation = config("orientation")
    
    ' -- Paper Size -- '
    ws.PageSetup.PaperSize = config("size")
    
    ' -- Scaling -- '
    ws.PageSetup.Zoom = config("scaling")
    
    ' -- Print Quiality -- '
    ws.PageSetup.PrintQuality = config("quality")
    
End Sub


'************************************************************
'Purpose:       Subrutine that sets an excel sheet's Margins
'Input:         Worksheet, active worksheet to be modify
'               Dictionary, dict with parameters to be set
'               Example: sectionMarginsSetup(activeWorksheet, DictProperties)
'Output:        Nothing
'************************************************************
Private Sub sectionMarginsSetup(ws As Worksheet, config As Dictionary)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("  SETTING UP MARGIN")
    
    ' -- Set Margins -- '
    ws.PageSetup.TopMargin = Application.InchesToPoints(config("top"))
    ws.PageSetup.BottomMargin = Application.InchesToPoints(config("bottom"))
    
    ws.PageSetup.LeftMargin = Application.InchesToPoints(config("left"))
    ws.PageSetup.RightMargin = Application.InchesToPoints(config("right"))
    
    ws.PageSetup.HeaderMargin = Application.InchesToPoints(config("top"))
    ws.PageSetup.FooterMargin = Application.InchesToPoints(config("top"))
    
End Sub


'************************************************************
'Purpose:       Subrutine that sets an excel sheet's Header & Footer Properties
'Input:         Worksheet, active worksheet to be modify
'               Dictionary, dict with parameters to be set
'               Example: sectionHeaderFooterSetup(activeWorksheet, DictProperties)
'Output:        Nothing
'************************************************************
Private Sub sectionHeaderFooterSetup(ws As Worksheet, config As Dictionary)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("  SETTING UP HEADER/FOOTER")
    
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


'************************************************************
'Purpose:       Subrutine that sets an excel sheet's Rows & Columns to be repeated.
'Input:         Worksheet, active worksheet to be modify
'               Example: sectionRepeatRowColSetup(activeWorksheet)
'Output:        Nothing
'************************************************************
Private Sub sectionRepeatRowColSetup(ws)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("  SETTING UP ROW/COLUMN REPEAT")
    
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


'************************************************************
'Purpose:       Subrutine that prepares an excel sheet's to be
'               ready for the creation of an Annex A. It does so by
'               setting up the columns to be repeated and caculating
'               some off set values.
'Input:         Worksheet, active worksheet to be modify
'               Example: SectionPrintArea(activeWorksheet)
'Output:        Nothing
'************************************************************
Private Sub SectionPrintArea(ws As Worksheet)

    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("  SETTING UP PRINT PREVIEW AREA")
    anxMessages.set_message ("      ***")
    
    Dim RepCols_width As Double
    Dim dimensions As New Dictionary
    
    'Wrap Text in the entire Doc
    anxMessages.set_message ("      Wrapping off/on Text")
    Range(GenerateColRowTitleRange(lastCl(ws), "Col")).WrapText = True
    
    'AutoFit all columns in Annex A
    anxMessages.set_message ("      AutoFit Columns")
    Columns(GenerateColRowTitleRange(lastCl(ws), "Col")).AutoFit
    
    'AutoFit all rows in Annex A
    anxMessages.set_message ("      AutoFit Rows")
    Rows(GenerateColRowTitleRange(CStr(lastRw(ws)), "Row")).AutoFit
    
    
    'Set up Reparing columns
    'setupPrintArea_RepCol ws, "A", annexA("repeating").item("Col")
    setupPrintArea_RepCol ws, "A", lastCl(ws)
    
    ' --- CALCULATIONS ---
    'Get width of RepCols, scale down & store  width of:
    '   Repeating Columns,
    '   Repeating Columns with left & right annexes borders, and
    '   Dynamic Columns
    anxMessages.set_message ("      Calculate Dynamic Columns Area")
    RepCols_width = CDbl(Format(Range(GenerateColRowTitleRange(annexA("repeating").item("Col"), "Col")).Width / Application.InchesToPoints(1), "0.00"))
    dimensions.Add "repCols", scaleDown(RepCols_width, CDbl(annexA("pagesetup")("scaling")))
    'dimensions.Add "repCols", RepCols_width
    dimensions.Add "repColsMrgs", Format((annexA("marginsetup")("left") * 2 + dimensions("repCols")), "0.00")
    dimensions.Add "dynCols", 11 - dimensions("repColsMrgs")
    ' --- CALCULATIONS ---
    
    'Set up Dynamic columns
    setupPrintArea_DynCol ws, annexA("repeating")("Col"), dimensions("dynCols")
    
    anxMessages.set_message ("      ***")
    
    'Activate AH Flag now tht data-headers have been arrange.
    setArrangeHeadersFlag (1)
    
End Sub


'************************************************************
'Purpose:       Subrutine that arranges the dynamic columns of an Annex A.
'Input:         Worksheet, active worksheet to be modify
'               String, column letter indication off set to start dynamic area.
'               Double, parameter that guides the space left to place dynamic columns.
'               Example: setupPrintArea_DynCol(activeWorksheet, "E", 8.25)
'Output:        Nothing
'************************************************************
Private Sub setupPrintArea_DynCol(ws As Worksheet, ByVal columnSetoff As String, y As Double)
    
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("      Method - PrintAra Dynamic Columns")
    
    ' FOR ANALYSIS - DELETE LATER
    ''    Dim fname As String
    ''    fname = "C:\Users\jairo\Desktop\Me\ExcelMacro\LiveRepo\Excel-Tools\Src\Analysis\" & Format(Now, "dd-mm-yyyy-hh-mm-ss") & ".txt"
    ''    Dim FSO As New FileSystemObject
    ''    Set FSO = CreateObject("Scripting.FileSystemObject")
    ''    Set FileToWrite = FSO.CreateTextFile(fname, ForWriting)
    ' FOR ANALYSIS - DELETE LATER
    
    
    Dim start, lmt As Integer
    
    start = Range((columnSetoff + ":" + columnSetoff)).Column + 1
    lmt = Range((lastCl(ws) + ":" + lastCl(ws))).Column
    
    Dim cw, tcw, delta, spreadDelta  As Double
    Dim numCols, colCounter, endSubset, strSubset As Integer
    
    tcw = 0#
    colCounter = start
    strSubset = start
    
    ' FOR ANALYSIS - DELETE LATER
    ''    FileToWrite.WriteLine "START OF setupPrintArea_DynCol"
    ''    FileToWrite.WriteLine "STRAT: " & start
    ''    FileToWrite.WriteLine "END: " & lmt
    ' FOR ANALYSIS - DELETE LATER

    For i = start To lmt
        
        'NOTE:
        '   TRYING NEW APPROACH - INSTEAD OF CONVERTING FROM POINTS TO INCHES, LET JUST HANDLE ALL THE MEASURMENT IN POINTS
        '   WHICH IS EXCEL'S DEFAULT UNIT OF MEASURMENT
        '   NEED TO REVISE: setupPrintArea_RepCol
        
        cw = Columns(i).Width / Application.InchesToPoints(1)
        cw = scaleDown(cw, 51.05)                'CDbl(annexA("pagesetup")("scaling")))
        tcw = tcw + cw
        
        ' FOR ANALYSIS - DELETE LATER
        ''        FileToWrite.WriteLine Columns(i).address & " | Widht: " & CStr(cw) & " | Total wc: " & CStr(tcw) & " | Length: " & CStr(y)
        ' FOR ANALYSIS - DELETE LATER
        
        If tcw > y Then
            tcw = Format(tcw - cw, "0.00")
            delta = Format(y - tcw, "0.00")
            spreadDelta = Format(delta / (i - colCounter), "0.00")
            
            ' FOR ANALYSIS - DELETE LATER
            ''            FileToWrite.WriteLine " tcw > y"
            ''            FileToWrite.WriteLine "i: " & CStr(i) & " - colCounter: " & CStr(colCounter) & " (i-colCounter): " & CStr(i - colCounter)
            ''            FileToWrite.WriteLine "**tcw: " & CStr(tcw)
            ''            FileToWrite.WriteLine "**Delta: " & CStr(delta)
            ''            FileToWrite.WriteLine "**spreadDelta: " & CStr(spreadDelta)
            ''            FileToWrite.WriteLine " ----------------------------------- "
            ' FOR ANALYSIS - DELETE LATER
            
            adjustWidths strSubset, i - 1, spreadDelta
            
            strSubset = i
            colCounter = i
            tcw = 0#
            delta = 0#
            spreadDelta = 0#
        End If
    Next i
    
    'FileToWrite.Close

End Sub


'************************************************************
'Purpose:       Subrutine that calculates the widht of a dynamic Column
'Input:         Integer, start of column position
'               Integer, start of next column
'               Double, difference between star and end of column
'               Example: adjustWidths(102,120,18.0)
'Output:        Nothing
'************************************************************
Private Sub adjustWidths(colStart As Integer, colEnd As Integer, delta As Double)
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("          Method - Adjust Widths")
    
    Dim iter As Long
    Dim newWidth As Double
    
    
    For i = colStart To colEnd
        ' NEWWIDTH IS IN POINTS
        newWidth = Columns(i).Width + delta
        'SETTING WIDTHS IN POINTS
        For iter = 1 To 5
            Columns(i).ColumnWidth = newWidth / Columns(i).Width * Columns(i).ColumnWidth
            ' Columns(i).ColumnWidth = Format(Columns(i).Width / Application.InchesToPoints(1), "0.00") + delta
        Next iter
    Next i
End Sub


'************************************************************
'Purpose:       Subrutine that sets the repeating Columns in an Annex A
'Input:         Worksheet, active worksheet to be modify
'               String, column letter indication the start position
'               String, column letter indication where repeation ends.
'               Example: setupPrintArea_RepCol(activeWorksheet, "A","E")
'Output:        Nothing
'************************************************************
Private Sub setupPrintArea_RepCol(ws As Worksheet, startCol As String, limitCol As String)

    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("      Method - PrintAra Repeating Columns")
    

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
        If (arrangeHeadersFlag = 0 And lenHeader > lenCell) Then
            'Arrange header width
            Range(headerRange).Value = arrangeHeader(Range(headerRange).Value)
        End If
        'Autofit current column
        Columns(Split(Cells(1, i).address, "$")(1)).AutoFit
        'Reset lenCell for next Column
        lenCell = 0
    Next

End Sub


'************************************************************
'Purpose:       Subrutine that initiates the printing of a dictionary
'Input:         Nothing
'               Example: printConfig()
'Output:        Nothing
'************************************************************
Public Sub printConfig()
    'STORE MARCRO LOGS/MESSAGES
    anxMessages.set_message ("Printing Annex A setup")
    
    printHashT annexA

End Sub


'************************************************************
'Purpose:       Subrutine that traverses a HashTable recrusively
'               printing all of its key:value attributes.
'Input:         Dictionary, Hashtable to traverse
'               Integer, level of deepness of the hashtable
'               Example: printHashT({k:v}, 0)
'Output:        Nothing
'************************************************************
Private Sub printHashT(dict As Dictionary, Optional ByVal level = 0)
    Dim indent As String
    indent = Space(level)
    
    For Each k In dict.keys
        If VarType(dict(k)) = 9 Then             ' vbObject == 9
            Debug.Print indent; k; ":  "
            level = level + 1
            printHashT dict(k), level
            level = level - 1
        Else
            Debug.Print indent; "-"; k; ": "; dict(k)
        End If
    Next
        
End Sub


'************************************************************
'Purpose:       Method that prints messages added to the message class.
'Input:         Nothing
'               Example: printLogs()
'Output:        Nothing
'************************************************************
Public Function printLogs() As String
    printLogs = anxMessages.get_messages
    anxMessages.clear_messages
End Function


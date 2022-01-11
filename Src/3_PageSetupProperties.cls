VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub pageSetup(ws As Worksheet)
    
    ' -- Orientation -- '
    ws.pageSetup.Orientation = xlLandscape
    
    ' -- Paper Size -- '
    ws.pageSetup.PaperSize = xlPaperLetter
    
    ' -- Scaling -- '
    ws.pageSetup.Zoom = 40
    
    ' -- Print Quiality -- '
    ws.pageSetup.PrintQuality = 600

End Sub

Sub marginSetup(ws As Worksheet)
    ' -- Marings -- '
    
    Dim topBottom As Double
    Dim leftRight As Double
    Dim headerFooter As Double
        
    topBottom = 0.5
    leftRight = 2.5
    headerFooter = 0.3
    
    ws.pageSetup.TopMargin = Application.InchesToPoints(topBottom)
    ws.pageSetup.BottomMargin = Application.InchesToPoints(topBottom)
    
    ws.pageSetup.LeftMargin = Application.InchesToPoints(topBottom)
    ws.pageSetup.RightMargin = Application.InchesToPoints(topBottom)
    
    ws.pageSetup.HeaderMargin = Application.InchesToPoints(headerFooter)
    ws.pageSetup.FooterMargin = Application.InchesToPoints(headerFooter)
    
End Sub



Sub headerfooterSetup(ws As Worksheet)

' -- Diff odd/even pages -- '
ws.pageSetup.OddAndEvenPagesHeaderFooter = False

' -- Diff 1st page -- '
ws.pageSetup.DifferentFirstPageHeaderFooter = False

' -- Scale with Doc -- '
ws.pageSetup.ScaleWithDocHeaderFooter = False

' -- Margin Aligment -- '
ws.pageSetup.AlignMarginsHeaderFooter = True

End Sub


Sub fileSetup()
    
    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate
    
    Call pageSetup(ws)
    Call marginSetup(ws)
    Call headerfooterSetup(ws)

End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PixelRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Private Declare PtrSafe Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
    Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
    Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
    Private Declare PtrSafe Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
#Else
    Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
    Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
    Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
    Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
#End If

Private Const LOGPIXELSY As Long = 90

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type
Public Function getLabelPixel(label As String) As Integer

  Dim font As New StdFont
  Dim sz As SIZE
  font.Name = "Arial Narrow"
  font.SIZE = 9.5

  sz = GetLabelSize(label, font)
  getLabelPixel = sz.cx

End Function

Private Function GetLabelSize(text As String, font As StdFont) As SIZE
    Dim tempDC As Long
    Dim tempBMP As Long
    Dim f As Long
    Dim lf As LOGFONT
    Dim textSize As SIZE

    ' Create a device context and a bitmap that can be used to store a
    ' temporary font object
    tempDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0)
    tempBMP = CreateCompatibleBitmap(tempDC, 1, 1)

    ' Assign the bitmap to the device context
    DeleteObject SelectObject(tempDC, tempBMP)

    ' Set up the LOGFONT structure and create the font
    lf.lfFaceName = font.Name & Chr$(0)
    lf.lfHeight = -MulDiv(font.SIZE, GetDeviceCaps(GetDC(0), 90), 72) 'LOGPIXELSY
    lf.lfItalic = font.Italic
    lf.lfStrikeOut = font.Strikethrough
    lf.lfUnderline = font.Underline
    If font.Bold Then lf.lfWeight = 800 Else lf.lfWeight = 400
    f = CreateFontIndirect(lf)

    ' Assign the font to the device context
    DeleteObject SelectObject(tempDC, f)

    ' Measure the text, and return it into the textSize SIZE structure
    GetTextExtentPoint32 tempDC, text, Len(text), textSize

    ' Clean up (very important to avoid memory leaks!)
    DeleteObject f
    DeleteObject tempBMP
    DeleteDC tempDC
  ' Return the measurements
    GetLabelSize = textSize

End Function

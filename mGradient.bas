Attribute VB_Name = "mGradient"
Option Explicit

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
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub FontStuff(picDraw As PictureBox, ByVal Caption As String, Optional ByVal Angulo As Integer = 90)
    
    On Error GoTo GetOut
    Dim f As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
    Dim FONTSIZE As Integer
    FONTSIZE = 10 'Val(txtSize.Text)
    
    f.lfEscapement = 10 * Angulo 'Val(txtDegree.Text) 'rotation angle, in tenths
    FontName = "Tahoma" + Chr$(0) 'null terminated
    f.lfFaceName = FontName
    f.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(f)
    hPrevFont = SelectObject(picDraw.hdc, hFont)
    
    picDraw.CurrentX = 3
    
    If Angulo <> 90 Then
        picDraw.CurrentY = 410
    End If
    
    picDraw.CurrentY = picDraw.Height - 10
    picDraw.Print Caption
    
    '  Clean up, restore original font
    hFont = SelectObject(picDraw.hdc, hPrevFont)
    DeleteObject hFont
    
    Exit Sub
GetOut:
    Exit Sub

End Sub


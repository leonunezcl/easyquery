Attribute VB_Name = "Strings"
Option Explicit
'Wert um 1 erhöhen
Public Function sys_Inc(value As Variant) As Variant
  sys_Inc = value + 1
End Function
'Wert um 1 vermindern
Public Function sys_Dec(value As Variant) As Variant
  sys_Dec = value - 1
End Function
'Wert1 und Wert2 vertauschen
Public Sub sys_Trade(value1, value2 As Variant)
Dim x As Variant
  
  x = value1
  value1 = value2
  value2 = x

End Sub
'leifert Wahr, wenn der angegebene String keine Zeichen enthält
Public Function sys_StrIsEmpty(str As Variant) As Boolean

  If Len(str) = 0 Then
    sys_StrIsEmpty = True
  Else
    sys_StrIsEmpty = False
  End If
    
End Function
'ist der übergebene Wert ein numerisches Zeichen
Public Function sys_IsNumChar(value As String) As Boolean
Dim x As Double

If sys_StrIsEmpty(value) Then Exit Function  'Parameter ungültig

Select Case value
  Case "0" To "9"
    sys_IsNumChar = True
  Case Else
    sys_IsNumChar = False
End Select
    
End Function
' Wandelt alle numerischen Werte des Strings in einen Intergerwert um
Public Function sys_StrToInt(str As String) As Integer
Dim i As Integer
Dim x As String
Dim NewStr As String

'Init
NewStr = ""

If sys_StrIsEmpty(str) Then ' String enthält keine Werte
  sys_StrToInt = 0
  Exit Function
End If

For i = 1 To Len(str)
  x = Mid(str, i, 1)
  If sys_IsNumChar(x) Then NewStr = NewStr + x
Next i
  
x = Val(NewStr)
sys_StrToInt = CInt(x)

End Function
'Die übergebene Ascii-Zeichenkette in hexadizimale Daten umwandeln
Public Function sys_StringToHex(str As String, isSeperator As Boolean, SeperatorChar As String, IsLineFeed As Boolean) As String
On Error GoTo Error

Dim x, y As String
Dim Result As String
Dim i As Integer

  If sys_StrIsEmpty(str) Then GoTo Error 'string enhält keine Daten
  If isSeperator And sys_StrIsEmpty(SeperatorChar) Then GoTo Error 'leeres seperatorzeichen ungültig
  If isSeperator And Not sys_StrScan(str, SeperatorChar) Then GoTo Error 'seperatorzeichen nicht im übergeben String enthalten
  
  
  For i = 1 To Len(str)
    x = Mid(str, i, 1)
    y = Hex(Asc(x))
    If Len(y) = 1 Then y = "0" + y
    
    If isSeperator Then
      Result = Result + y + SeperatorChar
    Else
      Result = Result + y
    End If
    
  Next i
  
  If IsLineFeed Then Result = Result + vbNewLine
  sys_StringToHex = Result
  
  Exit Function
  
Error:
  sys_StringToHex = ""
  Exit Function
  
End Function
'ist der übergebene Wert hexadizimal 0,1,2... - F
Public Function sys_isHexChar(str As String) As Boolean
  If sys_StrIsEmpty(str) Then Exit Function
  
  Select Case UCase(str)
    Case "0" To "9", "A" To "F"
      sys_isHexChar = True
    Case Else
      sys_isHexChar = False
  End Select
        
End Function
'wandelt einen hexadizimalen Wert in einen 32 Bit - Integerwert um
Public Function sys_HexToInt(HexStr As String) As Long
Dim x As Variant
  On Error GoTo PError
  
  If sys_StrIsEmpty(HexStr) Then Exit Function
  
  x = "&H" & HexStr
  sys_HexToInt = CDec(x)
  Exit Function

'Error Handler
PError:
  sys_HexToInt = 0
  Exit Function
End Function
'Wandelt eine Hexadizimale Zeichekette (8 Bit) in einen String um
'Der gebildete String enthält eine Datentiefe von 8-Bit
'die Daten können anschließend direkt an eine serielle schnittstelle
'gesendet werden
Public Function sys_HexToAsc(HexStr As Variant, isSeperator As Boolean, SeperatorChar As String, Optional KillCtrlChars As Boolean) As String
On Error GoTo Error
Dim Maxlen As Long
Dim Result, MyHexString, x As String
 
  If sys_StrIsEmpty(HexStr) Then GoTo Error
  If isSeperator And sys_StrIsEmpty(SeperatorChar) Then GoTo Error 'leeres seperatorzeichen ungültig
  If isSeperator And Not sys_StrScan(HexStr, SeperatorChar) Then GoTo Error 'seperatorzeichen nicht im übergeben String enthalten
      
  'init
  MyHexString = HexStr
  Maxlen = Len(MyHexString)
  
  While Len(MyHexString) >= 2
    If Not isSeperator Then 'ohne Verwendung eines Seperator-Zeichens immer zwei Zeichen verarbeiten
      Result = Result + Chr(sys_HexToInt(sys_StrCopy(MyHexString, 1, 2, True)))
    Else
      If Len(MyHexString) > 2 Then
        x = sys_StrGetTo(MyHexString, SeperatorChar, True)
        x = sys_StrSearchAndReplace(x, SeperatorChar, "")
      Else
        x = sys_StrCopy(MyHexString, 1, 2, True)
      End If
      
      Result = Result + Chr(sys_HexToInt(x))
    End If
  Wend
  
  If KillCtrlChars Then Result = sys_StrKillCtrlChars(Result)
  sys_HexToAsc = Result
  
  Exit Function
  
Error:
  sys_HexToAsc = ""
  Exit Function
 
End Function

'Wandelt hexadizimale Daten in einen String um. Hiebei werden alle nicht darstellbaren
'Zeichen herausgefiltert

Public Function sys_HexToStr(HexStr As Variant, isSeperator As Boolean, SeperatorChar As String) As String
Dim MyHexString As String
  MyHexString = sys_StrKillCrLf(HexStr)
  sys_HexToStr = sys_HexToAsc(MyHexString, isSeperator, SeperatorChar, True)
End Function

'Durchsucht einen String auf das erste Vorkommen und gibt die Position zurück
'Positon = 0 wenn nichts gefunden wurde
Public Function sys_StrPos(str, SearchStr As Variant) As Long
  
  ' Ungültige Werte in den Eingabestrings abfangen
  If sys_StrIsEmpty(str) Or sys_StrIsEmpty(SearchStr) Then Exit Function
  sys_StrPos = InStr(1, str, SearchStr, vbTextCompare)
  
End Function
'Ist die zu suchende Zeichenkette im dem übergebenen String enthalten
Public Function sys_StrScan(str, SearchStr As Variant) As Boolean
  sys_StrScan = sys_StrPos(str, SearchStr) <> 0
End Function
'an der übergeben Position einen Teilstring einfügen
Public Function sys_StrIns(str, InsStr As Variant, ByVal InsPos As Long) As String
Dim leftstr, Result, x As String

  If sys_StrIsEmpty(str) Then GoTo Error
  If sys_StrIsEmpty(InsStr) Then GoTo Error
  
  'Wenn Einfügeposition größer als Stringlänge ist, den Wert
  'hinten anhängen
  If InsPos > Len(str) Then
    sys_StrIns = str + InsStr
    Exit Function
  End If
  
  If InsPos <> 0 Then
    InsPos = sys_Dec(InsPos)
  End If
    
  x = str
  leftstr = sys_StrCopy(x, 1, InsPos, True)
  Result = leftstr + InsStr + x
     
  sys_StrIns = Result
  Exit Function
  
Error:
  sys_StrIns = ""
  Exit Function
  
End Function
'einen bestimmten Bereich aus einem String kopieren
'mit dem optionalen Flag [DeleteCopyChars], können die kopierten Zeichen aus dem String entfernt werden
Public Function sys_StrCopy(str As Variant, ByVal StartPos, copycount As Long, Optional DeleteCopyChars As Boolean) As String
Dim Result As String
  On Error GoTo Error
  
  ' Fehlerhafte Parameterangaben kompensieren
  If sys_StrIsEmpty(str) Then GoTo Error
  If StartPos > Len(str) Then GoTo Error
  If StartPos = 0 Then StartPos = 1
      
    
  If (StartPos - 1) + copycount > Len(str) Then
    copycount = Len(str) - StartPos
  End If
  
  Result = Mid(str, StartPos, copycount)
  
  If DeleteCopyChars Then
    str = sys_StrDel(str, StartPos, copycount)
  End If
  
  sys_StrCopy = Result
  
  Exit Function

'Fehlerbehandlung
Error:
  sys_StrCopy = ""
  
End Function
' Eine bestimme Zeichenkette aus einem Staring entfernen
Public Function sys_StrDel(str As Variant, ByVal StartPos, DelCount As Long) As String
Dim Result As String
Dim LStr, RStr As String
  On Error GoTo Error
  
  ' Fehlerhafte Parameterangaben kompensieren
  If sys_StrIsEmpty(str) Then GoTo Error
  If StartPos > Len(str) Then GoTo Error
  If StartPos <= 1 Then
    StartPos = 0
  Else
    StartPos = StartPos - 1
  End If
      
  If StartPos + DelCount > Len(str) Then
    DelCount = Len(str) - StartPos
  End If
  
  If StartPos > 0 Then
    LStr = Mid(str, 1, StartPos)
  Else
    LStr = ""
  End If
  
  RStr = Mid(str, (StartPos + DelCount) + 1, Len(str) - (StartPos + DelCount))
  sys_StrDel = LStr + RStr
  
  Exit Function
  
'Fehlerbehandlung
Error:
  sys_StrDel = ""
  Exit Function
 
End Function
'Teilstring bis Position [ToStr] auslesen
Public Function sys_StrGetTo(str, ToStr As Variant, DeleteCopyChars As Boolean) As String
Dim Result As String
 
  If sys_StrIsEmpty(ToStr) Or Not sys_StrScan(str, ToStr) Then
    Exit Function
  End If

  Result = sys_StrCopy(str, 1, sys_StrPos(str, ToStr), DeleteCopyChars)
  sys_StrGetTo = Result
  
End Function
'Teilstring ab Position [FromStr] auslesen
Public Function sys_StrGetFrom(str, FromStr As Variant, DeleteCopyChars As Boolean) As String
Dim Result As String
Dim pos, length, x As Long

  If sys_StrIsEmpty(FromStr) Or Not sys_StrScan(str, FromStr) Then
    Exit Function
  End If

  pos = sys_StrPos(str, FromStr)
  length = Len(str)
  
  'INIT
  If pos < length Then
    sys_Inc (pos)
    x = 1
  Else
    x = 0
  End If
  
  Result = sys_StrCopy(str, pos, Abs(pos - length) + x, DeleteCopyChars)
  sys_StrGetFrom = Result
  
End Function

'Suche und ersetze. Wenn die gefundene Vorkommnisse entfernt werden
'sollen, muß der Parameter ReplaceStr leer sein ("")
Public Function sys_StrSearchAndReplace(str, SearchStr, ReplaceStr As Variant) As String
Dim FoundPos As Long
Dim Result As String


On Error GoTo Error

  'übergeben parameter prüfen
  If sys_StrIsEmpty(str) Then GoTo Error
  If sys_StrIsEmpty(SearchStr) Then GoTo Error 'Suchzeichen ungültig
  If Not sys_StrScan(str, SearchStr) Then GoTo Error 'Suchzeichen nicht vorhanden
    
  'init
  Result = str
  
  While sys_StrPos(Result, SearchStr) <> 0
    FoundPos = sys_StrPos(Result, SearchStr)
    
    If Not sys_StrIsEmpty(ReplaceStr) Then
      Result = sys_StrIns(Result, ReplaceStr, FoundPos)
      Result = sys_StrDel(Result, FoundPos + 1, 1)
    Else
      Result = sys_StrDel(Result, FoundPos, 1)
    End If
  Wend
  
  sys_StrSearchAndReplace = Result
  Exit Function
  
Error:
  sys_StrSearchAndReplace = ""
  Exit Function
  
End Function
'einen Teilstring (von Zeichen bis Zeichen) aus dem übergebenen String herausziehen
Public Function sys_StrGetArea(str, StartChar, EndChar As Variant, Optional DeleteArea As Boolean) As String
Dim p1, p2 As Long
Dim StrResult As String

On Error GoTo Error
  
  'Fehlerhafte Parameter abfangen
  'Start und Endezeichen müssen auf jeden Fall vorhanden sein
  'Die übergebene Masterzeichkette muß Werte aufweisen
  If sys_StrIsEmpty(str) Or (Not sys_StrScan(str, StartChar) Or Not _
    sys_StrScan(str, EndChar)) Then GoTo Error
    
  p1 = sys_StrPos(str, StartChar)
  p2 = sys_StrPos(str, EndChar)
  If p1 > p2 Then Call sys_Trade(p1, p2)
    
  StrResult = sys_StrCopy(str, p1, p2, DeleteArea)
  sys_StrGetArea = StrResult
    
  Exit Function
  
Error:
  Exit Function

End Function
'Die Steuerzeichen Dezimal 13 und 10 aus dem String entfernen
Public Function sys_StrKillCrLf(str As Variant) As String
Dim i As Long
Dim mystring, Result, x As String
  
  If sys_StrIsEmpty(str) Then Exit Function
  mystring = str
  
  For i = 1 To Len(mystring)
    x = sys_StrCopy(mystring, i, 1)
    If Asc(x) <> 13 And Asc(x) <> 10 Then Result = Result + x
  Next
  
  sys_StrKillCrLf = Result
End Function
'ist das übergebene Zeichen ein darstellbares Ascii-Zeichen
Public Function sys_StrIsNormChar(char As String) As Boolean
Dim Result As Boolean
Dim x As Long

  If sys_StrIsEmpty(char) Then Exit Function
  x = Asc(char)
  
  Select Case x
    Case 10, 13
      Result = True
      
    Case 32 To 126
      Result = True
      
    Case Else
      Result = False
  End Select
  
  sys_StrIsNormChar = Result
End Function
'Alle Seuerzeichen aus dem String entfernen
'CrLf bleibt erhalten
Public Function sys_StrKillCtrlChars(str As Variant) As String
Dim i, CharCode As Long
Dim mystring, Result, x As String

  If sys_StrIsEmpty(str) Then Exit Function
    mystring = str
  
    For i = 1 To Len(mystring)
      x = sys_StrCopy(mystring, i, 1)
      If sys_StrIsNormChar(x) Then Result = Result + x
    Next
    
    sys_StrKillCtrlChars = Result
End Function

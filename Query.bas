Attribute VB_Name = "mEasyQuery"
Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public gbInicio As Boolean
Public glbFontLetra As String
Public glbFontSize As Variant
Public glbBackColor As Variant
Public glbColorSql As Variant
Private SqlHis As String
Public glbNumSQl As Variant
Public glbSalir As Boolean

Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EMPTYUNDOBUFFER = &HCD

Private Const LF_FACESIZE = 32
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CYCAPTION = 4
Public Const SQL_SUCCESS As Long = 0
Public Const SQL_FETCH_NEXT As Long = 1
Public Const RGN_OR = 2
Public Const C_INI = "EQUERY.INI"
Public Const WM_UNDO = &H304
Public Const WM_USER = &H400
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Private Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
Private Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
Private Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
Private Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.

'Common Dialog
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000

Private Const SWP_NOZORDER = &H4
Public Const GW_HWNDPREV = 3
Private Const CSIDL_PERSONAL = &H5

Public Cargando As Boolean
Public glbSQl As String
Public glbTimeOut
Public TrxActiva As Boolean
Public Conectado As Boolean
Public gsBuffer As String
Public glbMisDocumentos As String

Public Const OF_EXIST = &H4000
Public Const OFS_MAXPATHNAME = 256

Public Type ECharrange
    cpMin As Long
    cpMax As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PointAPI
    x As Long
    y As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

'Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
Public Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)
Public Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Public Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
'Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Dim PrevProc As Long

Type eSentencias
    Color As Long
    Glosa As String
End Type
Public aSentencias() As eSentencias
Public CustomColors() As Byte

Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, _
                                    ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&

Private Type PageSetupDlg
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As PointAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type ChooseFont
    lStructSize As Long
    hWndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
                                   '    contains cust. dlg. template
    lpszStyle As String          '  return the style field here
                                   '  must be LF_FACESIZE or bigger
    nFontType As Integer          '  same value reported to the EnumFonts
                                   '    call back with the extra FONTTYPE_
                                   '    bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
                                   '    CF_LIMITSIZE is used
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Enum ShowCommands
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
End Enum

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "COMDLG32.DLL" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PageSetupDlg Lib "COMDLG32.DLL" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type

Public Inicio As Boolean

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const IDC_WAIT = 32514&
Private Const IDC_ARROW = 32512&
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private lRet As Long

Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CUT = &H300
Public Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Public Sub CargaHistorial(lvw As ListView)

    Dim nFreeFile As Long
    Dim Linea As String
    Dim K As Integer
    Dim ArrTemp() As String
    Dim c As Integer
    
    If Len(glbMisDocumentos) > 0 Then
        nFreeFile = FreeFile
        K = 1
        lvw.ListItems.Clear
        
        ReDim ArrTemp(0)
        
        If VBOpenFile(glbMisDocumentos & "easyquery.his") Then
            Open glbMisDocumentos & "easyquery.his" For Input Shared As #nFreeFile
                Do While Not EOF(nFreeFile)
                    Line Input #nFreeFile, Linea
                    ReDim Preserve ArrTemp(K)
                    ArrTemp(K) = Linea
                    K = K + 1
                Loop
            Close #nFreeFile
            
            'grabar de atras para adelante
            c = 1
            For K = UBound(ArrTemp) To 1 Step -1
                lvw.ListItems.Add , "k" & c, Format(c, "0000"), 4, 4
                lvw.ListItems("k" & c).SubItems(1) = ArrTemp(K)
                c = c + 1
            Next K
        End If
    End If
    
End Sub

Public Function GetSpecialfolder() As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim Path As String
    
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, CSIDL_PERSONAL, IDL)
    If r = 0 Then
        'Create a buffer
        Path$ = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        Path = Left$(Path, InStr(Path, Chr$(0)) - 1)
        
        If Right$(Path, 1) <> "\" Then
            Path = Path & "\"
        End If
        
        GetSpecialfolder = Path
                
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

'grabar el historial de sqls digitados
Public Function GrabaHistorialSQL() As Boolean

    On Local Error GoTo ErrorGrabaHistorialSQL
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    Dim K As Integer
    Dim Nuevo As Boolean
    
    ret = True
    
    nFreeFile = FreeFile
    
    If Len(glbMisDocumentos) > 0 Then
        K = 1
        'verificar si existe archivo
        If VBOpenFile(glbMisDocumentos & "easyquery.his") Then
            Open glbMisDocumentos & "easyquery.his" For Input Shared As #nFreeFile
                Do While Not EOF(nFreeFile)
                    Line Input #nFreeFile, SqlHis
                    K = K + 1
                Loop
            Close #nFreeFile
        End If
        
        Nuevo = False
        If K > glbNumSQl Then
            Nuevo = True
        End If
        
        SqlHis = frmMain.ActiveForm.txtQuery.Text
        SqlHis = Replace(SqlHis, Chr$(10), " ")
        SqlHis = Replace(SqlHis, Chr$(13), " ")
                
        If Nuevo Then
            If Len(SqlHis) > 0 Then
                Open glbMisDocumentos & "easyquery.his" For Output As #nFreeFile
                    Print #nFreeFile, SqlHis
                Close #nFreeFile
            End If
        Else
            If Len(SqlHis) > 0 Then
                Open glbMisDocumentos & "easyquery.his" For Append As #nFreeFile
                    Print #nFreeFile, SqlHis
                Close #nFreeFile
            End If
        End If
    End If
    
    GoTo SalirGrabaHistorialSQL
    
ErrorGrabaHistorialSQL:
    ret = False
    MsgBox "GrabaHistorialSQL : " & Err & " " & Error$, vbCritical
    Resume SalirGrabaHistorialSQL
    
SalirGrabaHistorialSQL:
    GrabaHistorialSQL = ret
    Err = 0
    
End Function


Public Function VBOpenFile(ByVal Archivo As String) As Boolean

    Dim ret As Boolean
    Dim lRet As Long
    Dim of As OFSTRUCT
    
    ret = False
    
    lRet = OpenFile(Archivo, of, OF_EXIST)
    
    If of.nErrCode = 0 Then ret = True
    
    VBOpenFile = ret
    
End Function
Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Function ConexionActiva(ByVal Conexion As String) As Integer

    Dim K As Integer
    
    Dim ret As Integer
    
    For K = 1 To UBound(cState)
        If cState(K).Deleted = False Then
            If cState(K).Conexion = Conexion Then
                ret = K
                Exit For
            ElseIf cState(K).LlaveMdb = Conexion Then
                ret = K
                Exit For
            End If
        End If
    Next K
    
    ConexionActiva = ret
    
End Function

Public Sub Deshacer(ByVal hwnd As Long)

    Dim ret As Long
    
    ret = SendMessage(hwnd, WM_UNDO, 0, 0)
    
End Sub

Public Sub ColorSQL(rtb As Control, ByVal sSearch As String, ByVal Color As ColorConstants)

    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim sql As String
        
    lPos = 1
        
    sql = UCase$(rtb.Text)
    
    Do While lPos < Len(sql)
        
        'get sub string from the text
        'this is because the InStr() returns the
        'position of first occurence of the string...
        sTmp = Mid(sql, lPos, Len(sql))
        
        'find the string in sub string
        lWhere = InStr(sTmp, UCase$(sSearch))
        'accumulate the lPos to be relative to the actual text
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            rtb.SelStart = lPos - 2   ' set selection start and
            rtb.SelLength = Len(sSearch)   ' set selection length.   Else
            rtb.SelColor = Color
            rtb.SelBold = True
            rtb.SelLength = 0
            rtb.SelBold = False
        Else
            Exit Do 'we are ready
        End If
    Loop
    
End Sub

Public Function ShowColor(hwnd As Long) As Long
    
    Dim cc As ChooseColor
    Dim Custcolor(16) As Long
    'Dim lReturn As Long
    'Dim CustomColors() As Byte
    
    ReDim CustomColors(0)
    
    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hWndOwner = hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If ChooseColor(cc) <> 0 Then
        ShowColor = cc.rgbResult
        'CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
    
    'Erase CustomColors
    
End Function
Sub CenterWindow(ByVal hwnd As Long)

    Dim wRect As RECT
    
    Dim x As Integer
    Dim y As Integer

    Dim ret As Long
    
    ret = GetWindowRect(hwnd, wRect)
    
    x = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hwnd, vbNull, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub
Sub ActivatePrevInstance()

   Dim OldTitle As String
   Dim PrevHndl As Long
   Dim result As Long

   'Save the title of the application.
   OldTitle = App.Title

   'Rename the title of this application so FindWindow
   'will not find this application instance.
   App.Title = "unwanted instance"

   'Attempt to get window handle using VB4 class name.
   PrevHndl = FindWindow("ThunderRTfrmMain", OldTitle)

   'Check for no success.
   If PrevHndl = 0 Then

      'Attempt to get window handle using VB5 class name.
      PrevHndl = FindWindow("ThunderRT5frmMain", OldTitle)
   End If

   'Check if found
   If PrevHndl = 0 Then
        'Attempt to get window handle using VB6 class name
        PrevHndl = FindWindow("ThunderRT6frmMain", OldTitle)
   End If

   'Check if found
   If PrevHndl = 0 Then
      'No previous instance found.
      Exit Sub
   End If

   'Get handle to previous window.
   PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

   'Restore the program.
   result = OpenIcon(PrevHndl)

   'Activate the application.
   result = SetForegroundWindow(PrevHndl)

   Call ShowWindow(PrevHndl, SW_SHOWMAXIMIZED)
    
   'End the application.
   End

End Sub
Public Sub Main()

    Cargando = True
    
    If App.PrevInstance Then ActivatePrevInstance: Exit Sub
    
    glbFontLetra = LeeIni("Editor", "Letra", C_INI)
    If glbFontLetra = "" Then glbFontLetra = "Verdana"
    
    glbFontSize = LeeIni("Editor", "Size", C_INI)
    If glbFontSize = "" Then glbFontSize = 10
    
    glbBackColor = LeeIni("Editor", "BackColor", C_INI)
    If glbBackColor = "" Then glbBackColor = QBColor(15)
    
    glbColorSql = LeeIni("Editor", "ColorFuente", C_INI)
    If glbColorSql = "" Then glbColorSql = QBColor(1)
    
    glbMisDocumentos = LeeIni("Historial", "Path", C_INI)
    If glbMisDocumentos = "" Then glbMisDocumentos = GetSpecialfolder()
    
    glbNumSQl = LeeIni("Historial", "Numero", C_INI)
    If glbNumSQl = "" Then glbNumSQl = 100
    
    frmAcercaDe.Show
            
End Sub

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ComboBox) As Long
    
    Dim FaceName As String
    'Dim FullName As String
    
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
    
End Function



Public Function ShowPrinter(frmOwner As Form, Optional PrintFlags As Long) As Integer
    
    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = frmOwner.hwnd

    PrintDlg.flags = PrintFlags
    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    
    ShowPrinter = False
    
    If PrintDialog(PrintDlg) <> 0 Then
        
        ShowPrinter = True
        
        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    End If
End Function


Public Function ShowPageSetupDlg(ByVal hwnd As Long) As Long
    
    Dim m_PSD As PageSetupDlg
    'Set the structure size
    m_PSD.lStructSize = Len(m_PSD)
    'Set the owner window
    m_PSD.hWndOwner = hwnd
    'Set the application instance
    m_PSD.hInstance = App.hInstance
    'no extra flags
    m_PSD.flags = 0

    'Show the pagesetup dialog
    If PageSetupDlg(m_PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
    
End Function

Public Sub GrabaIni(ByVal ArchivoIni As String, ByVal Seccion As String, ByVal Llave As String, ByVal Valor)

    Dim ret As Long
    
    ret = WritePrivateProfileString(Seccion, Llave, CStr(Valor), ArchivoIni)
    
End Sub

Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function

Public Sub AnimaWindow(ByVal hwnd As Long, ByVal Tiempo As Long)

    AnimateWindow hwnd, Tiempo, AW_CENTER Or AW_VER_POSITIVE Or AW_HOR_NEGATIVE Or AW_HIDE
    
End Sub


Public Sub Copiar()
        
    Clipboard.Clear
    Clipboard.SetText frmMain.ActiveForm!txtQuery.SelText
                
End Sub

Public Sub Cortar(ByVal hwnd As Long)

    Dim ret As Long
    
    ret = SendMessageByVal(hwnd, WM_CUT, 0, 0)
    
End Sub

Public Function SaveDialog(ByVal hwnd As Long, Filter As String, Title As String, InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    Dim A As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = hwnd
    ofn.hInstance = App.hInstance
    
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    A = GetSaveFileName(ofn)

    If (A) Then
        SaveDialog = Trim$(ofn.lpstrFile)
    Else
        SaveDialog = ""
    End If

End Function






Public Sub Hourglass(hwnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hwnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If
    DoEvents
    
End Sub


Public Function ClearEnterInString(ByVal sText As String) As String

    Dim K As Integer
    
    Dim ret As String
    
    For K = 1 To Len(sText)
        If Chr$(Asc(Mid$(sText, K, 1))) <> Chr$(13) And Chr$(Asc(Mid$(sText, K, 1))) <> Chr$(10) Then
            ret = ret & Mid$(sText, K, 1)
        Else
            ret = ret & " "
        End If
    Next K
    
    ClearEnterInString = ret
    
End Function


Public Function LeeIni(ByVal Seccion As String, ByVal Llave As String, ByVal ArchivoIni As String) As String

    Dim lRet As Long
    Dim ret As String
    
    Dim buffer As String
    
    buffer = String$(255, " ")
    
    lRet = GetPrivateProfileString(Seccion, Llave, "", buffer, Len(buffer), ArchivoIni)
    
    buffer = Trim$(buffer)
    ret = Left$(buffer, Len(buffer) - 1)
    
    LeeIni = ret
    
End Function



Public Sub Pegar(ByVal hwnd As Long)

    Dim ret As Long
    
    frmMain.ActiveForm.txtQuery.Text = ""
    ret = SendMessage(hwnd, WM_PASTE, 0, 0)
        
End Sub


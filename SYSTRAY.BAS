Attribute VB_Name = "mSysTray"
Option Explicit

'-------------------------------------------------------
' Api Declares....
'-------------------------------------------------------
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

'-------------------------------------------------------
' Api Constants...
'-------------------------------------------------------
Public Const GWL_USERDATA = (-21&)
'Public Const GWL_WNDPROC = (-4&)
Public Const WM_USER = &H400&
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&

'Public Const WM_MOUSEMOVE = &H200&
'Public Const WM_LBUTTONDOWN = &H201&
'Public Const WM_LBUTTONUP = &H202&
'Public Const WM_LBUTTONDBLCLK = &H203&
'Public Const WM_RBUTTONDOWN = &H204&
'Public Const WM_RBUTTONUP = &H205&
'Public Const WM_RBUTTONDBLCLK = &H206&

'DrawEdge constants
Public Const BDR_RAISEDOUTER = &H1&
Public Const BDR_RAISEDINNER = &H4&
Public Const BF_LEFT = &H1&             ' Border flags
Public Const BF_TOP = &H2&
Public Const BF_RIGHT = &H4&
Public Const BF_BOTTOM = &H8&
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const BF_SOFT = &H1000&          ' For softer buttons

'-------------------------------------------------------
' Api Types....
'-------------------------------------------------------
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

Public PrevWndProc As Long

'------------------------------------------------------------
Public Function SubWndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------------------
' This is the control subclassed window proc.
'------------------------------------------------------------
    Dim SysTray As cSysTray                         ' SysTray class variable
    Dim ClassAddr As Long                           ' long pointer to class object
'------------------------------------------------------------
    
    'If hwnd = frmMain.hwnd Then
        Select Case msg                                 ' Determine
            'Case WM_CLOSE, WM_DESTORY
                'SubWndProc = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
            '    SysTray.InTray = False
            '    Exit Function
            Case TRAY_CALLBACK                              ' Callback message received when user clicks on system tray...
                ' Retrieve long pointer to class object, this was saved in the _
                  USERDATA of the window struct. after the user control was subclassed...
                ClassAddr = GetWindowLong(hwnd, GWL_USERDATA) ' get pointer to object
                CopyMemory SysTray, ClassAddr, 4            ' Copy an unreferenced pointer to object into variable
                
                SysTray.SendEvent lParam, wParam            ' Send windows message\user event to control
                
                CopyMemory SysTray, 0&, 4                   ' Nullify object pointer
        End Select
        
        ' Forward all messages to previous window procedure...(This must be done)
        SubWndProc = CallWindowProc(PrevWndProc, hwnd, msg, wParam, lParam)
    'End If
    
'------------------------------------------------------------
End Function
'------------------------------------------------------------


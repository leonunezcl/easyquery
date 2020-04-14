Attribute VB_Name = "MAppActivate"
  Option Explicit
  
  ' A demonstration of finding a window  and activating it
  ' using API functions by
  ' Bryan Stafford of New Vision Software® Copyright © 1998
  ' This code is released into the public domain "AS IS"
  ' without warranty of any kind.  In other words, use at
  ' your own risk.
  
  ' This example will either look for an application with a variable title (like
  ' Notepad or Word that displays the current document name in the titlebar) or
  ' just find an application that has a static title depending on how you set
  ' this constant:
  
  Public Const VARIABLE_TITLE As Boolean = True
  
  ' compile the app and then try to run multipule instances.  Also try minimizing
  ' the current instance and then try to start another instance.
  
  ' IMPORTANT NOTE: this code will only work after it is compiled!  It will NOT
  ' work in the VB IDE.
  

  ' API declares, constants and types
  
  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type  ' RECT

  Public Const SW_RESTORE As Long = 9&

  Public Const GW_CHILD As Long = 5&
  Public Const GW_HWNDNEXT As Long = 2&

  Public Declare Function GetDesktopWindow& Lib "user32" ()
  
  Public Declare Function GetWindow& Lib "user32" (ByVal hWnd&, ByVal wCmd&)

  Public Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" _
                                    (ByVal hWnd&, ByVal lpString$, ByVal cch&)
  
  Public Declare Function ShowWindow& Lib "user32" (ByVal hWnd&, ByVal nCmdShow&)

  Public Declare Function GetWindowRect& Lib "user32" (ByVal hWnd&, lpRect As RECT)

  Public Declare Function MoveWindow& Lib "user32" (ByVal hWnd&, ByVal x&, _
                          ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal bRepaint&)

  Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
  
  Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
                                          (ByVal lpClassName$, ByVal lpWindowName$)

Public Sub ActivatePrevInstance(ByVal hAppWindow&)

  ' tell the previous instance to restore itself incase it's minimized
  Call ShowWindow(hAppWindow, SW_RESTORE)
  
  ' get it's current size and change it to the initial size and position
  Dim theRect As RECT
  GetWindowRect hAppWindow, theRect
  MoveWindow hAppWindow, ((Screen.Width \ Screen.TwipsPerPixelX) \ 2) - ((theRect.Right - theRect.Left) \ 2), _
        ((Screen.Height \ Screen.TwipsPerPixelY) \ 3) - ((theRect.Bottom - theRect.Top) \ 3), _
        theRect.Right - theRect.Left, theRect.Bottom - theRect.Top, True
       
  ' activate the thread for the prievious instance
  Call SetForegroundWindow(hAppWindow)

End Sub

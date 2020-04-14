VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   3855
   ClientTop       =   1665
   ClientWidth     =   4560
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMainSkin 
      AutoSize        =   -1  'True
      Height          =   1755
      Left            =   0
      Picture         =   "Inicio.frx":0BC2
      ScaleHeight     =   1695
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterWindow hWnd
    
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
     
    Dim WindowRegion As Long
    
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW)
End Sub



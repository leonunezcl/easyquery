VERSION 5.00
Begin VB.Form frmStop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EasyQuery !"
   ClientHeight    =   1110
   ClientLeft      =   2835
   ClientTop       =   4245
   ClientWidth     =   2655
   ControlBox      =   0   'False
   Icon            =   "frmStop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStop 
      Caption         =   "Detener carga de registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   30
      Picture         =   "frmStop.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   2550
   End
   Begin VB.Label lblRegistros 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2445
      TabIndex        =   1
      Top             =   855
      Width           =   135
   End
End
Attribute VB_Name = "frmStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStop_Click()
    frmMain.m_Detener = True
End Sub

Private Sub Form_Load()
    CenterWindow hwnd
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmStop = Nothing
End Sub



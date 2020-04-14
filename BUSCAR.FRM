VERSION 5.00
Object = "{D4E70ACD-6152-11D3-81C5-8ECDDCEC9A77}#1.0#0"; "AXANIM.OCX"
Object = "{89A6DB9C-F3B9-11D4-A70B-0000E82185CB}#1.0#0"; "XtremeProgressBar.ocx"
Begin VB.Form frmBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutando Consulta ..."
   ClientHeight    =   1425
   ClientLeft      =   4845
   ClientTop       =   3075
   ClientWidth     =   4290
   ControlBox      =   0   'False
   Icon            =   "Buscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin XtremeProgresBar.XtremeProgressBar pgb 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   661
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Verdana"
      FontSize        =   8,25
      Rojo            =   0
      Verde           =   0
   End
   Begin axAnimation.Animation anim 
      Height          =   900
      Left            =   90
      Top             =   60
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   1588
      ResourceID      =   160
      BorderStyle     =   0
      Appearance      =   0
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterWindow hwnd
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Set frmBuscar = Nothing
    
End Sub



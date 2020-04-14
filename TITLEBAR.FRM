VERSION 5.00
Begin VB.Form frmAcercaDe 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de QueryIt !"
   ClientHeight    =   6120
   ClientLeft      =   1560
   ClientTop       =   4065
   ClientWidth     =   8610
   Icon            =   "Titlebar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   405
      Left            =   3960
      TabIndex        =   0
      Top             =   5610
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creado por Luis ""JEDI"" Núñez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   870
      TabIndex        =   3
      Top             =   1110
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Query It !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Query It !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   0
      Picture         =   "Titlebar.frx":030A
      Top             =   0
      Width           =   8640
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdAceptar_Click()
    
    Me.Hide
    DoEvents
    
    Unload Me
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not gbInicio Then
        Main.Show
        gbInicio = True
    End If
    
    Set frmAcercaDe = Nothing
    
End Sub



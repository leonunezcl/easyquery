VERSION 5.00
Begin VB.Form frmParametro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros"
   ClientHeight    =   2625
   ClientLeft      =   2085
   ClientTop       =   2085
   ClientWidth     =   6765
   Icon            =   "frmParametro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3840
      TabIndex        =   12
      Top             =   2205
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Datos del Parámetro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   6630
      Begin VB.TextBox txtDefecto 
         Height          =   285
         Left            =   1590
         TabIndex        =   4
         Top             =   1635
         Width           =   4875
      End
      Begin VB.TextBox txtLargo 
         Height          =   285
         Left            =   1590
         TabIndex        =   3
         Top             =   1305
         Width           =   915
      End
      Begin VB.ComboBox cboIO 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   945
         Width           =   4890
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   570
         Width           =   4890
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1590
         TabIndex        =   0
         Top             =   225
         Width           =   4890
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Valor x Defecto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "I/O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   975
         Width           =   300
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   615
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   255
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Nuevo As Boolean
Public Nombre As String
Public tipo As ADODB.DataTypeEnum
Public Io As ADODB.ParameterDirectionEnum
Public largo As Long
Public Defecto As String

Private Sub Form_Load()

    Call CenterWindow(hwnd)
    
    Call Hourglass(hwnd, True)
        
    If Nuevo Then
        Me.Caption = "Nuevo Parametro"
    Else
        Me.Caption = "Modificar Parametro"
        txtNombre.Text = Nombre
                
        txtLargo.Text = largo
        txtDefecto.Text = Defecto
    End If
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmParametro = Nothing
End Sub



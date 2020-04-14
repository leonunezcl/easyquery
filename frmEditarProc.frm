VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNuevoProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Procedimiento"
   ClientHeight    =   4560
   ClientLeft      =   3420
   ClientTop       =   3480
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "frmEditarProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1440
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2550
      Width           =   7110
   End
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
      Index           =   4
      Left            =   5655
      TabIndex        =   6
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
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
      Index           =   3
      Left            =   4320
      TabIndex        =   5
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Eliminar"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Modificar"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Agregar"
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
      Left            =   345
      TabIndex        =   2
      Top             =   4110
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwParam 
      Height          =   1545
      Left            =   105
      TabIndex        =   1
      Top             =   705
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   2725
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "I/O"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tamaño"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Valor x Defecto"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtNombre 
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Sentencia SQL"
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
      TabIndex        =   9
      Top             =   2325
      Width           =   1290
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Parámetros"
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
      TabIndex        =   8
      Top             =   450
      Width           =   960
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
      TabIndex        =   7
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmNuevoProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IndexConexion As Integer
Public Procedimiento As String

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
    
    ElseIf Index = 1 Then
    
    ElseIf Index = 2 Then
    
    ElseIf Index = 3 Then
    
    ElseIf Index = 4 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hwnd, True)
    
    Call CenterWindow(hwnd)
    
    If Procedimiento <> "" Then
    
    End If
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmNuevoProc = Nothing
End Sub



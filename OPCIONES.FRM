VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   4080
   ClientLeft      =   1365
   ClientTop       =   2745
   ClientWidth     =   7200
   Icon            =   "Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   5880
      TabIndex        =   9
      Top             =   3675
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comandos críticos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   6825
      Begin VB.TextBox txtComando 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2445
      End
      Begin VB.ListBox lisComando 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   930
         Width           =   2445
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estos comando criticos se ejecutaran solo con la confirmacion por parte del usuario."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2625
         TabIndex        =   8
         Top             =   2385
         Width           =   3885
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "UPDATE , DELETE , ALTER, DROP ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2625
         TabIndex        =   7
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplos de comandos criticos :"
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
         Left            =   2625
         TabIndex        =   6
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label Label2 
         Caption         =   "Marque la casilla para activar comando critico. Para desactivar desmarque comando critico."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   2625
         TabIndex        =   5
         Top             =   600
         Width           =   4050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digite comando ..."
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip tabOpciones 
      Height          =   3525
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6218
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguridad"
            Object.ToolTipText     =   "Configura comandos criticos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()

    Dim NComandos
    Dim k As Integer
    
    ReDim Comandos(0)
    
    NComandos = lisComando.ListCount - 1
         
    For k = 0 To NComandos
        ReDim Preserve Comandos(k + 1)
        Comandos(k + 1).Comando = lisComando.List(k)
        Comandos(k + 1).Activo = lisComando.Selected(k)
    Next k
        
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim NComandos
    Dim k As Integer
    
    NComandos = UBound(Comandos)
         
    CenterWindow hwnd
    
    For k = 1 To NComandos
        lisComando.AddItem Comandos(k).Comando
        If Comandos(k).Activo Then
           lisComando.Selected(k - 1) = True
        End If
    Next k
    
End Sub

Private Sub txtComando_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim$(txtComando.Text) <> "" Then
            lisComando.AddItem txtComando.Text
            lisComando.Selected(lisComando.NewIndex) = True
        End If
    End If
    
End Sub



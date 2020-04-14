VERSION 5.00
Begin VB.Form frmFind 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   1215
   ClientLeft      =   1245
   ClientTop       =   4095
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FIND.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   150
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   4380
      TabIndex        =   4
      ToolTipText     =   "Salir de la pantalla"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Buscar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4380
      TabIndex        =   3
      ToolTipText     =   "Buscar la palabra digitada"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "&Solo palabras completas"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   2805
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "&Según Modelo"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   660
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    gbMatchCase = chkMatchCase.Value
    gbWholeWord = chkWholeWord.Value
    gsFindText = txtFind.Text
    Call FindText
End Sub

Private Sub Form_Load()
    
    CenterWindow hwnd
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
    
    chkMatchCase.Value = gbMatchCase
    chkWholeWord.Value = gbWholeWord
    txtFind.Text = gsFindText
    txtFind.SelLength = Len(gsFindText)
    
    glbFindSql = ClearEnterInString(frmMain.ActiveForm.txtQuery.Text)
    gsFindText = glbFindSql
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmFind = Nothing
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = (txtFind.Text <> "")
End Sub


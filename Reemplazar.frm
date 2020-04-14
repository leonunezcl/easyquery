VERSION 5.00
Begin VB.Form frmReemplazar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reemplazar ...."
   ClientHeight    =   1515
   ClientLeft      =   2445
   ClientTop       =   4215
   ClientWidth     =   7230
   Icon            =   "Reemplazar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   7230
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
      Height          =   405
      Index           =   2
      Left            =   5730
      TabIndex        =   4
      Top             =   990
      Width           =   1425
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Reemplazar Todo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   5730
      TabIndex        =   3
      Top             =   540
      Width           =   1425
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Reemplazar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   5730
      TabIndex        =   2
      Top             =   90
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reemplazar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtReemplazar 
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Top             =   720
         Width           =   4035
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Top             =   360
         Width           =   4035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reemplazar por"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   390
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmReemplazar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frm As New frmQuery
Private Sub Reemplazar(ByVal Tipo As Integer)

    On Local Error GoTo ErrorReemplazar
    
    Dim sBuscar As String
    Dim sReemplazar As String
    Dim lPos As Integer
    
    sBuscar = Trim$(txtBuscar.Text)
    sReemplazar = Trim$(txtReemplazar.Text)
    
    If sBuscar = "" Then Exit Sub
    
    gsFindText = sBuscar
    
    If Tipo = 1 Then
        If gbLastPos = 0 Or gbLastPos > Len(gsBuffer) Then
            lPos = 1
        Else
            lPos = gbLastPos
        End If
    Else
        lPos = 1
        gsBuffer = frmMain.ActiveForm!txtQuery.Text
    End If
    
    If Tipo = 1 Then
        lPos = InStr(lPos, UCase$(gsBuffer), UCase$(sBuscar))
        If lPos <> 0 Then
            gbLastPos = lPos
            Mid(gsBuffer, lPos, Len(sReemplazar)) = sReemplazar
        Else
            gbLastPos = 0
        End If
    Else
        gsBuffer = Replace(gsBuffer, sBuscar, sReemplazar)
        gbLastPos = 0
    End If
    
    frmMain.ActiveForm!txtQuery.Text = ""
    frmMain.ActiveForm!txtQuery.SelStart = 1
    frmMain.ActiveForm!txtQuery.SelColor = glbColorSql
    frmMain.ActiveForm!txtQuery.SelLength = Len(gsBuffer)
    frmMain.ActiveForm!txtQuery.Text = gsBuffer
    
    frmMain.ActiveForm.HayCambios = True
    'frm.txtQuery.Visible = False
    Call frmMain.ActiveForm.FormateaSentencias
    'frm.txtQuery.Visible = True
    
    Exit Sub
    
ErrorReemplazar:
    Err = 0
    gbLastPos = 0
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'R
            Call Reemplazar(1)
        Case 1  'R.t
            Call Reemplazar(2)
        Case 2  'Salir
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
    
    txtBuscar.Text = gsFindText
    
    gsBuffer = UCase$(frmMain.ActiveForm!txtQuery.Text)
        
    gbLastPos = 0
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmReemplazar = Nothing
    
End Sub



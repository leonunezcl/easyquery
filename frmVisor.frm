VERSION 5.00
Begin VB.Form frmVisor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor"
   ClientHeight    =   3900
   ClientLeft      =   3810
   ClientTop       =   3240
   ClientWidth     =   6615
   Icon            =   "frmVisor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2550
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Height          =   3360
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   6510
   End
End
Attribute VB_Name = "frmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Call CenterWindow(hwnd)
            
    txtValor.Text = frmMain.ActiveForm.lviewCampos.SelectedItem.SubItems(1)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmVisor = Nothing
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form frmBrowserTablas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browser de Tablas"
   ClientHeight    =   5715
   ClientLeft      =   2595
   ClientTop       =   2085
   ClientWidth     =   3360
   Icon            =   "Tablas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.SkinForm SkinForm 
      Height          =   480
      Left            =   1050
      OleObjectBlob   =   "Tablas.frx":08CA
      TabIndex        =   0
      Top             =   2040
      Width           =   480
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3030
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tablas.frx":091B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treTablas 
      Height          =   5385
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   9499
      _Version        =   393217
      Indentation     =   229
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione tabla :"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "frmBrowserTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim rs As ADODB.Recordset
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    
    SkinForm.SkinPath = glbPathSkin
        
    Call RemueveMenu(hwnd)
    
    treTablas.Nodes.Clear
    treTablas.Nodes.Add(, , "EJ1", "Tablas", 1).EnsureVisible
    
    For Each rs In DBConnection
        treTablas.Nodes.Add("EJ1", tvwChild, rs.Name, rs.Name, 1).EnsureVisible
    Next
    
    Set rs = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmTablas = Nothing
    
End Sub



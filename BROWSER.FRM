VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browser"
   ClientHeight    =   4755
   ClientLeft      =   3900
   ClientTop       =   2430
   ClientWidth     =   2670
   Icon            =   "Browser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3090
      Top             =   4080
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
            Picture         =   "Browser.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treQuerys 
      Height          =   4305
      Left            =   60
      TabIndex        =   0
      Top             =   330
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   7594
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
      Caption         =   "Seleccione query :"
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
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   1605
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
        
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    
    treQuerys.Nodes.Clear
    treQuerys.Nodes.Add(, , "EJ1", "Querys", 1).EnsureVisible
    
    Dim ArrayCount As Integer
    Dim i As Integer
    
    ArrayCount = UBound(fState)

    For i = 1 To ArrayCount
        If Not fState(i).Deleted Then
            treQuerys.Nodes.Add("EJ1", tvwChild, "q" & i, Document(i).Caption, 1).EnsureVisible
        End If
    Next
    
    'SkinForm.SkinPath = glbPathSkin
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmMain.mnuOpciones_Browser.Checked = False
    
    Set frmBrowser = Nothing
    
End Sub


Private Sub treQuerys_NodeClick(ByVal Node As msComctlLib.Node)

    Dim k As Integer
    
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    For k = 1 To ArrayCount
        If Not fState(k).Deleted Then
            If fState(k).Tag = Node.Key Then
                Document(k).ZOrder 0
                Exit For
            End If
        End If
    Next
    
End Sub



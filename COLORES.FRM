VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmColores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Barra de Formato Texto"
   ClientHeight    =   465
   ClientLeft      =   1440
   ClientTop       =   3555
   ClientWidth     =   9660
   Icon            =   "Colores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdBold"
            Object.ToolTipText     =   "Negrita"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdItalic"
            Object.ToolTipText     =   "Italica"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdUnderline"
            Object.ToolTipText     =   "Subrayado"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   1140
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   3315
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   5820
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   45
         Width           =   255
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   1065
      End
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   600
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Colores.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Colores.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Colores.frx":0AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Colores.frx":0C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Colores.frx":0D12
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFont_Click()

    If Me.Visible Then
        If cboFont.ListIndex <> -1 Then
            frmMain.ActiveForm.txtQuery.Font = cboFont.List(cboFont.ListIndex)
        End If
    End If
    
End Sub

Private Sub cboSize_Click()

    If Me.Visible Then
        If cboSize.ListIndex <> -1 Then
            frmMain.ActiveForm.txtQuery.Font.Size = cboSize.List(cboSize.ListIndex)
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim k As Integer
    
    picColor(1).BackColor = QBColor(1)
    
    Call FillListWithFonts(cboFont)
    
    cboSize.AddItem "8"
    cboSize.AddItem "10"
    cboSize.AddItem "12"
    cboSize.AddItem "14"
    cboSize.AddItem "18"
    cboSize.AddItem "24"
    
    For k = 2 To 15
        Load picColor(k)
        picColor(k).Height = picColor(1).Height
        picColor(k).Width = picColor(1).Width
        picColor(k).Top = picColor(1).Top
        picColor(k).BackColor = QBColor(k)
        picColor(k).Left = picColor(k - 1).Left + 250
        picColor(k).Visible = True
    Next k
    
    'SkinForm.SkinPath = glbPathSkin
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmMain.mnuOpciones_FTexto.Checked = False
    'frmMain.Toolbar.Buttons("cmdFonts").Value = tbrUnpressed
    
    Set frmColores = Nothing
        
End Sub

Private Sub picColor_Click(Index As Integer)
    frmMain.ActiveForm.txtQuery.SelColor = picColor(Index).BackColor
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        Case "cmdBold"
            If Button.Value = tbrPressed Then
                frmMain.ActiveForm.txtQuery.SelBold = True
            Else
                frmMain.ActiveForm.txtQuery.SelBold = False
            End If
        Case "cmdItalic"
            If Button.Value = tbrPressed Then
                frmMain.ActiveForm.txtQuery.SelItalic = True
            Else
                frmMain.ActiveForm.txtQuery.SelItalic = False
            End If
        Case "cmdUnderline"
            If Button.Value = tbrPressed Then
                frmMain.ActiveForm.txtQuery.SelUnderline = True
            Else
                frmMain.ActiveForm.txtQuery.SelUnderline = False
            End If
    End Select
    
End Sub



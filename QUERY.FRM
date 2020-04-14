VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{3559B97F-0351-11D5-B855-444553540001}#1.0#0"; "XProgress.ocx"
Begin VB.Form frmQuery 
   HelpContextID   =  10
   AutoRedraw      =   -1  'True
   Caption         =   "Query : 1"
   ClientHeight    =   4650
   ClientLeft      =   -525
   ClientTop       =   2970
   ClientWidth     =   11940
   Icon            =   "Query.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Tag             =   """q1"""
   Begin Progress.XProgress pgbQuery 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSComctlLib.ImageCombo imgConexiones 
      Height          =   330
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   "Conexiones activas"
      Top             =   90
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgLst"
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm 
      Height          =   480
      Left            =   990
      OleObjectBlob   =   "Query.frx":08CA
      TabIndex        =   6
      Top             =   60
      Width           =   480
   End
   Begin FPSpreadADO.fpSpread griQuery 
      Height          =   2085
      Left            =   3960
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   7725
      _Version        =   196608
      _ExtentX        =   13626
      _ExtentY        =   3678
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      Protect         =   0   'False
      SelectBlockOptions=   1
      SpreadDesigner  =   "Query.frx":09A1
      UserResize      =   1
   End
   Begin ComctlLib.ListView lviewDetalle 
      Height          =   2145
      Left            =   3930
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   3784
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nombre Físico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Largo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Acepta Nulos"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin RichTextLib.RichTextBox txtQuery 
      Height          =   2145
      Left            =   3930
      TabIndex        =   0
      Top             =   510
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   3784
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Query.frx":0BEF
   End
   Begin ComctlLib.ListView lviewCampos 
      Height          =   2415
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.StatusBar staQuery 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4365
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6615
            MinWidth        =   6615
            Object.Tag             =   ""
            Object.ToolTipText     =   "QueryIt !"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14111
            MinWidth        =   14111
            Object.Tag             =   ""
            Object.ToolTipText     =   "Listo"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabQuery 
      Height          =   2745
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   4842
      ImageList       =   "imgList"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sql"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Sql a ejecutar"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Resultado"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Resultado de la Búsqueda"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Detalle Campos"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
      MouseIcon       =   "Query.frx":0C9D
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Query.frx":0FB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Query.frx":47D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Query.frx":492F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   150
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Query.frx":4A8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Query.frx":4DA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Query.frx":50BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Query.frx":53D9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Detener As Boolean
Public m_Ejecutando As Boolean

Public Rs As New ADODB.Recordset

Private nContador As Integer
Public HayCambios As Boolean
Public m_Conexion As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 And Not Shift Then
        Call frmMain.EjecutaQuery
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'If KeyAscii = vbKeyF5 Then
    '    Call frmMain.EjecutaQuery
    'End If
    
End Sub

Private Sub Form_Load()
'*** Code added by HelpWriter ***
    SetApphelp Me.hWnd
'***********************************

    On Local Error Resume Next
    
    SkinForm.SkinPath = glbPathSkin

    nContador = 0
    
    HayCambios = True
    
    Call FormateaSentencias
    
    Err = 0
    
End Sub

Public Sub FormateaSentencias()

    Dim Sql As String
    Dim ret As Long
    Dim K As Integer
    
    Dim CHARRANGE As ECharrange
    
    Sql = Trim$(ClearEnterInString(txtQuery.Text))
    
    If Sql = "" Or Not HayCambios Then Exit Sub
        
    ret = SendMessage(txtQuery.hWnd, EM_EXGETSEL, 0, CHARRANGE)
    
    For K = 1 To UBound(aSentencias)
        Call ColorSQL(txtQuery, aSentencias(K).Glosa, aSentencias(K).Color)
    Next K
    
    HayCambios = False
        
    txtQuery.SelStart = CHARRANGE.cpMin
    txtQuery.SelLength = 0
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call frmMain.EliminarHojaConsulta
    
End Sub

Private Sub Form_Resize()
    
    On Local Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        TabQuery.Height = Me.ScaleHeight - staQuery.Height - 150
        TabQuery.Width = Me.ScaleWidth - lviewCampos.Width - 150
        txtQuery.Height = TabQuery.Height - 500
        txtQuery.Width = TabQuery.Width - 180
        griQuery.Height = txtQuery.Height
        griQuery.Width = txtQuery.Width
        lviewDetalle.Height = griQuery.Height
        lviewDetalle.Width = griQuery.Width
        lviewCampos.Height = Me.ScaleHeight - staQuery.Height - 440
        pgbQuery.Top = Me.ScaleHeight - 260
    End If
    
    Err = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
'*** Code added by HelpWriter ***
    QuitHelp
'***********************************

    On Local Error Resume Next
        
    fState(Val(Mid$(Me.Tag, 2))).Deleted = True
    
    Call AnimaWindow(hWnd, 50)
    
    DoEvents
    
    Set frmQuery = Nothing
    
    Err = 0
    
End Sub


Private Sub griQuery_Click(ByVal Col As Long, ByVal Row As Long)

    Dim c As Integer
    Dim K As Integer
    Dim itmx As ListItem
    Dim Valor
    
    If lviewCampos.ListItems.Count > 0 And Row > 0 Then
        For K = 2 To griQuery.MaxCols
            Call griQuery.GetText(K, Row, Valor)
            Set itmx = lviewCampos.ListItems(K - 1)
            itmx.SubItems(1) = Valor
        Next K
    End If
    
End Sub

Private Sub griQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If frmMain.mnuQuery.Enabled Then
            PopupMenu frmMain.mnuQuery
        End If
    End If
    
End Sub


Private Sub imgConexiones_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Private Sub lviewDetalle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If lviewDetalle.ListItems.Count > 0 Then
            PopupMenu frmMain.mnuTablas
        End If
    End If
    
End Sub


Private Sub TabQuery_Click()
    
    If TabQuery.SelectedItem.Index = 1 Then
        txtQuery.ZOrder 0
        txtQuery.Visible = True
        griQuery.Visible = False
        lviewDetalle.Visible = False
    ElseIf TabQuery.SelectedItem.Index = 2 Then
        griQuery.ZOrder 0
        griQuery.Visible = True
        txtQuery.Visible = False
        lviewDetalle.Visible = False
    Else
        lviewDetalle.ZOrder 0
        lviewDetalle.Visible = True
        griQuery.Visible = False
        txtQuery.Visible = False
    End If
    
End Sub

Private Sub Timer1_Timer()

    If nContador = -1 Then Exit Sub
    
    nContador = nContador + 1
    
    If nContador = 2 Then
        Call FormateaSentencias
        nContador = -1
    End If
    
End Sub

Private Sub txtQuery_GotFocus()
    Timer1.Enabled = True
End Sub


Private Sub txtQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    nContador = 0
End Sub


Private Sub txtQuery_KeyPress(KeyAscii As Integer)

    Dim ret As Integer
    Dim ret2 As Long
    Dim CHARRANGE As ECharrange
    
    If KeyAscii <> vbKeyReturn Then
        ret = SendMessage(txtQuery.hWnd, EM_EXGETSEL, 0, CHARRANGE)
        txtQuery.SelStart = CHARRANGE.cpMin
        txtQuery.SelLength = 0
        txtQuery.SelColor = 0
    End If
        
    nContador = 0
    HayCambios = True
    
End Sub


Private Sub txtQuery_LostFocus()

    Call FormateaSentencias
    HayCambios = False
    nContador = -1
    Timer1.Enabled = False
    
End Sub


Private Sub txtQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu frmMain.mnuEdicion
    End If
    
End Sub


Private Sub XProgress1_Resize()

End Sub



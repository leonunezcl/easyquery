VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPreview 
   Caption         =   "Presentación Preliminar"
   ClientHeight    =   5715
   ClientLeft      =   645
   ClientTop       =   1755
   ClientWidth     =   10830
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   10830
   Begin VB.ComboBox cboPorcentaje 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   90
      Width           =   1245
   End
   Begin VB.ComboBox cboVerPagina 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   2295
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Imprimir"
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
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin FPSpreadADO.fpSpreadPreview fpPreview 
      Height          =   5625
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   7065
      _Version        =   196608
      _ExtentX        =   12462
      _ExtentY        =   9922
      _StockProps     =   96
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   1
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Porcentaje"
      Height          =   195
      Left            =   4680
      TabIndex        =   4
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ver Página"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lHwnd As Long
Public Estructura As Boolean
Private Sub cboPorcentaje_Click()
    fpPreview.PageViewPercentage = cboPorcentaje.List(cboPorcentaje.ListIndex)
End Sub

Private Sub cboVerPagina_Click()

    fpPreview.PageViewType = cboVerPagina.ListIndex
    
    If cboVerPagina.ListIndex = 2 Then
        cboPorcentaje.Enabled = True
    Else
        cboPorcentaje.Enabled = False
    End If
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    Select Case Index
        Case 0
            frmImprimir.Show vbModal
        Case 1
        
    End Select
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    fpPreview.hWndSpread = lHwnd '
    
    cboVerPagina.AddItem "Página Entera"
    cboVerPagina.AddItem "Tamaño Normal"
    cboVerPagina.AddItem "Porcentaje"
    cboVerPagina.AddItem "Ancho Página"
    cboVerPagina.AddItem "Largo Página"
    'cboVerPagina.AddItem "Múltiples Páginas"
    
    cboVerPagina.ListIndex = 0
    
    For k = 10 To 100 Step 10
        cboPorcentaje.AddItem CStr(k)
    Next k
    
End Sub


Private Sub Form_Resize()

    fpPreview.Left = 0
    fpPreview.Height = ScaleHeight
    fpPreview.Width = ScaleWidth
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmPreview = Nothing
    
End Sub



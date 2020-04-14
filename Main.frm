VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQuery 
   Caption         =   "Query : 1"
   ClientHeight    =   3255
   ClientLeft      =   1515
   ClientTop       =   4740
   ClientWidth     =   12060
   ForeColor       =   &H8000000D&
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Tag             =   """q1"""
   Begin VB.Frame fraBus 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   135
      TabIndex        =   17
      Top             =   1380
      Width           =   3510
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   810
         TabIndex        =   19
         Text            =   "*.*"
         ToolTipText     =   "Puede usar marcadores. Ejemplo  Cli* , A* , Vi?A"
         Top             =   45
         Width           =   2565
      End
      Begin VB.Label lblBuscar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
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
         Left            =   105
         TabIndex        =   18
         Top             =   75
         Width           =   660
      End
   End
   Begin RichTextLib.RichTextBox txtQuery 
      Height          =   2145
      Left            =   3945
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   525
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   3784
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      MousePointer    =   3
      TextRTF         =   $"Main.frx":014A
   End
   Begin MSComctlLib.ListView lviewCampos 
      Height          =   1140
      Left            =   135
      TabIndex        =   4
      Top             =   1815
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabInfo 
      Height          =   1995
      Left            =   90
      TabIndex        =   11
      Top             =   1005
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3519
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vista &Rápida"
            Object.ToolTipText     =   "Visualizar la información del registro seleccionado"
            ImageVarType    =   2
            ImageIndex      =   9
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tablas"
            Object.ToolTipText     =   "Tablas de la conexión"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Vistas"
            Object.ToolTipText     =   "Vistas de la conexión"
            ImageVarType    =   2
            ImageIndex      =   8
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Procedimientos"
            Object.ToolTipText     =   "Procedimientos almacenados"
            ImageVarType    =   2
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
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
      Height          =   2775
      Left            =   3750
      MouseIcon       =   "Main.frx":01F8
      MousePointer    =   99  'Custom
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   9
      Tag             =   "0"
      Top             =   90
      Width           =   90
   End
   Begin MSComctlLib.StatusBar staQuery 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6438
            MinWidth        =   6438
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5745
      Top             =   1335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":04A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0602
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0FEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbQuery 
      Height          =   240
      Left            =   5280
      TabIndex        =   3
      Top             =   1245
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageCombo imgConexiones 
      Height          =   330
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Conexiones activas"
      Top             =   330
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      ImageList       =   "imgLst"
   End
   Begin FPSpreadADO.fpSpread griQuery 
      Height          =   2085
      Left            =   3960
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   7725
      _Version        =   196608
      _ExtentX        =   13626
      _ExtentY        =   3678
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AutoCalc        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   1
      MaxRows         =   1
      Protect         =   0   'False
      SpreadDesigner  =   "Main.frx":1146
      UserResize      =   1
      TextTip         =   4
      TextTipDelay    =   250
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   4080
      Top             =   2610
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
            Picture         =   "Main.frx":1395
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4BB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4D0D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabQuery 
      Height          =   2775
      Left            =   3840
      TabIndex        =   5
      Top             =   105
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   4895
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&SQL"
            Object.ToolTipText     =   "Digite intrucción sql"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Resultado búsqueda"
            Object.ToolTipText     =   "Resultado del sql escrito"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Información de campos"
            Object.ToolTipText     =   "Información de los campos de la consulta"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Historial de SQL"
            Object.ToolTipText     =   "Historial con los SQL ejecutados"
            ImageVarType    =   2
            ImageIndex      =   4
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
   Begin MSComctlLib.ListView lviewDetalle 
      Height          =   2070
      Left            =   3945
      TabIndex        =   7
      Top             =   540
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   3651
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre Físico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Largo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Decimales"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Acepta Nulos"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistorial 
      Height          =   2145
      Left            =   3945
      TabIndex        =   8
      Top             =   510
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sql"
         Object.Width           =   12347
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTablas 
      Height          =   1620
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Doble clic para ejecutar tabla seleccionada"
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre tabla"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwVistas 
      Height          =   1620
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Doble clic para ejecutar vista seleccionada"
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre Vista"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwProcs 
      Height          =   1620
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Doble clic para ejecutar procedimiento seleccionado"
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre procedimiento"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtHist 
      Height          =   2145
      Left            =   4095
      TabIndex        =   16
      Top             =   720
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   3784
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":4E69
   End
   Begin VB.Label lblInfoCon 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Información de la conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   12
      Top             =   705
      Width           =   3645
   End
   Begin VB.Label lblOrigenes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Seleccionar origen de datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   90
      Width           =   3645
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nContador As Integer
Public HayCambios As Boolean
Public m_Conexion As Integer
Private cat As New ADOX.Catalog
Private cmdx As New ADODB.Command
Private prm As New ADODB.Parameter
Private Rs As ADODB.Recordset
Private itmx As ListItem
Private flag As Boolean

' flag to indicate that a splitter recieved a mousedown
Private fInitiateDrag As Boolean
Private Const SPLT_WDTH As Integer = 4
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long)


'actualizar informacion de procedimientos almacenados
Public Sub ActualizarProcedimientos()

    On Local Error GoTo ErrorActualizarProcedimientos
    
    Dim ci As ComboItem
    Dim ActiveConexion As Integer
    Dim K As Integer
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
        Call frmMain.HabilitaMenues(False)
        Call frmMain.HabilitaBotones(False)
        
        Load frmInfo
        frmInfo.Label1.Caption = "Actualizando información de procedimientos ..."
        frmInfo.Show
        DoEvents
        
        lvwProcs.ListItems.Clear
        
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
            
        'obtener info de la conexion
        'cargar procs
        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaProcedures)
        K = 1
        Do While Not Rs.EOF
            ReDim Preserve cState(ActiveConexion).aProcs(K)
            cState(ActiveConexion).aProcs(K).Nombre = Trim$(Rs!procedure_name)
            
            If IsNull(Rs!Description) Then
                cState(ActiveConexion).aProcs(K).Descripción = ""
            Else
                cState(ActiveConexion).aProcs(K).Descripción = Trim$(Rs!Description)
            End If
                
            Rs.MoveNext
            K = K + 1
        Loop
        
        Rs.Close
        
        'cargar procedimientos
        For K = 1 To UBound(cState(ActiveConexion).aProcs)
            Set itmx = lvwProcs.ListItems.Add(, "p" & K, cState(ActiveConexion).aProcs(K).Nombre, 7, 7)
            itmx.SubItems(1) = cState(ActiveConexion).aProcs(K).Descripción
        Next K
        
        tabInfo.Tabs(4).Caption = "&Procedimientos (" & lvwProcs.ListItems.Count & ")"
            
        Unload frmInfo
        Set ci = Nothing
        
        Call frmMain.HabilitaMenues(True)
        Call frmMain.HabilitaBotones(True)
        Call frmMain.HabiBotones2
        Call Hourglass(hwnd, False)
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
    Exit Sub
    
ErrorActualizarProcedimientos:
    Unload frmInfo
    MsgBox "ActualizarProcedimientos : " & Err & " " & Error$, vbCritical
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

'actualizar tablas
Public Sub ActualizarTablas()

    On Local Error GoTo ErrorActualizarTablas
    
    Dim ci As ComboItem
    Dim ActiveConexion As Integer
    Dim K As Integer
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
        Call frmMain.HabilitaMenues(False)
        Call frmMain.HabilitaBotones(False)
        
        Load frmInfo
        frmInfo.Label1.Caption = "Actualizando información de tablas ..."
        frmInfo.Show
        DoEvents
        
        lvwTablas.ListItems.Clear
        
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
            
        'obtener info de la conexion
        'cargar tablas
        Set Rs = New ADODB.Recordset
        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        
        ReDim cState(ActiveConexion).aTablas(0)
        
        K = 1
        Do While Not Rs.EOF
            If Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
                ReDim Preserve cState(ActiveConexion).aTablas(K)
                cState(ActiveConexion).aTablas(K).Nombre = Trim$(Rs!table_name)
                If IsNull(Rs!Description) Then
                    cState(ActiveConexion).aTablas(K).Descripción = ""
                Else
                    cState(ActiveConexion).aTablas(K).Descripción = Trim$(Rs!Description)
                End If
                K = K + 1
            End If
            Rs.MoveNext
        Loop
        
        Rs.Close
        
        'cargar tablas
        For K = 1 To UBound(cState(ActiveConexion).aTablas)
            Set itmx = lvwTablas.ListItems.Add(, "t" & K, cState(ActiveConexion).aTablas(K).Nombre, 5, 5)
            itmx.SubItems(1) = cState(ActiveConexion).aTablas(K).Descripción
        Next K
        
        tabInfo.Tabs(2).Caption = "&Tablas (" & lvwTablas.ListItems.Count & ")"
            
        Unload frmInfo
        Set ci = Nothing
        
        Call frmMain.HabilitaMenues(True)
        Call frmMain.HabilitaBotones(True)
        Call frmMain.HabiBotones2
        Call Hourglass(hwnd, False)
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
    Exit Sub
    
ErrorActualizarTablas:
    Unload frmInfo
    MsgBox "ActualizarTablas : " & Err & " " & Error$, vbCritical
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

'actualizar las vistas de la conexión
Public Sub ActualizarVistas()

    On Local Error GoTo ErrorActualizarVistas
    
    Dim ci As ComboItem
    Dim ActiveConexion As Integer
    Dim K As Integer
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
        Call frmMain.HabilitaMenues(False)
        Call frmMain.HabilitaBotones(False)
        
        Load frmInfo
        frmInfo.Label1.Caption = "Actualizando información de vistas ..."
        frmInfo.Show
        DoEvents
        
        lvwVistas.ListItems.Clear
        
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
            
        'cargar vistas
        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "VIEW"))
        K = 1
        Do While Not Rs.EOF
            If Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
                ReDim Preserve cState(ActiveConexion).aVistas(K)
                cState(ActiveConexion).aVistas(K).Nombre = Trim$(Rs!table_name)
                If IsNull(Rs!Description) Then
                    cState(ActiveConexion).aVistas(K).Descripción = ""
                Else
                    cState(ActiveConexion).aVistas(K).Descripción = Trim$(Rs!Description)
                End If
                K = K + 1
            End If
            Rs.MoveNext
        Loop
        
        Rs.Close
        
        'cargar vistas
        For K = 1 To UBound(cState(ActiveConexion).aVistas)
            Set itmx = lvwVistas.ListItems.Add(, "v" & K, cState(ActiveConexion).aVistas(K).Nombre, 8, 8)
            itmx.SubItems(1) = cState(ActiveConexion).aVistas(K).Descripción
        Next K
        
        tabInfo.Tabs(3).Caption = "&Vistas (" & lvwVistas.ListItems.Count & ")"
            
        Unload frmInfo
        Set ci = Nothing
        
        Call frmMain.HabilitaMenues(True)
        Call frmMain.HabilitaBotones(True)
        Call frmMain.HabiBotones2
        Call Hourglass(hwnd, False)
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
    Exit Sub
    
ErrorActualizarVistas:
    Unload frmInfo
    MsgBox "ActualizarVistas : " & Err & " " & Error$, vbCritical
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub


'agrega un procedimiento almacenado
Public Sub AddProc()

    Dim ActiveConexion As Integer
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        ActiveConexion = ConexionActiva(imgConexiones.SelectedItem.Text)
        frmNuevoProc.IndexConexion = ActiveConexion
        frmNuevoProc.Show vbModal
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub BuscarTabla()

    Dim K As Integer
    Dim Busqueda As String
        
    Dim ci As ComboItem
    Dim itmx As ListItem
    Dim ActiveConexion As Integer
    Dim j As Integer
    
    Call frmMain.HabilitaMenues(False)
    Call frmMain.HabilitaBotones(False)
    DoEvents
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
                
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
        
        On Local Error GoTo Salir
        
        'cargar tablas
        lvwTablas.ListItems.Clear
        
        Busqueda = Trim$(txtBuscar.Text)
        
        If Len(Trim$(Busqueda)) = 0 Or Busqueda = "*.*" Then
            'cargar todas
            For K = 1 To UBound(cState(ActiveConexion).aTablas)
                Set itmx = lvwTablas.ListItems.Add(, , cState(ActiveConexion).aTablas(K).Nombre, 5, 5)
                itmx.SubItems(1) = cState(ActiveConexion).aTablas(K).Descripción
            Next K
        Else
            j = 1
            For K = 1 To UBound(cState(ActiveConexion).aTablas)
                If cState(ActiveConexion).aTablas(K).Nombre Like Busqueda Then
                    Set itmx = lvwTablas.ListItems.Add(, , cState(ActiveConexion).aTablas(K).Nombre, 5, 5)
                    itmx.SubItems(1) = cState(ActiveConexion).aTablas(K).Descripción
                    j = j + 1
                End If
            Next K
            
            If lvwTablas.ListItems.Count = 0 Then
                MsgBox "No se han encontrado tablas con esa coincidencia.", vbCritical
            End If
        End If
        
        Call Hourglass(hwnd, False)
    End If
    
    Set ci = Nothing
    Set itmx = Nothing
    
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    
    Exit Sub:
Salir:
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub BuscarVista()

    Dim K As Integer
    Dim Busqueda As String
        
    Dim ci As ComboItem
    Dim itmx As ListItem
    Dim ActiveConexion As Integer
    Dim j As Integer
    
    Call frmMain.HabilitaMenues(False)
    Call frmMain.HabilitaBotones(False)
    DoEvents
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
                
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
        
        On Local Error GoTo Salir
        
        'cargar tablas
        lvwVistas.ListItems.Clear
        
        Busqueda = Trim$(txtBuscar.Text)
        
        If Len(Trim$(Busqueda)) = 0 Or Busqueda = "*.*" Then
            'cargar todas
            For K = 1 To UBound(cState(ActiveConexion).aVistas)
                Set itmx = lvwVistas.ListItems.Add(, , cState(ActiveConexion).aVistas(K).Nombre, 5, 5)
                itmx.SubItems(1) = cState(ActiveConexion).aVistas(K).Descripción
            Next K
        Else
            j = 1
            For K = 1 To UBound(cState(ActiveConexion).aVistas)
                If cState(ActiveConexion).aVistas(K).Nombre Like Busqueda Then
                    Set itmx = lvwVistas.ListItems.Add(, , cState(ActiveConexion).aVistas(K).Nombre, 5, 5)
                    itmx.SubItems(1) = cState(ActiveConexion).aVistas(K).Descripción
                    j = j + 1
                End If
            Next K
            
            If lvwVistas.ListItems.Count = 0 Then
                MsgBox "No se han encontrado vistas con esa coincidencia.", vbCritical
            End If
        End If
        
        Call Hourglass(hwnd, False)
    End If
    
    Set ci = Nothing
    Set itmx = Nothing
    
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    
    Exit Sub:
Salir:
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub BuscarProcedimiento()

    Dim K As Integer
    Dim Busqueda As String
        
    Dim ci As ComboItem
    Dim itmx As ListItem
    Dim ActiveConexion As Integer
    Dim j As Integer
    
    Call frmMain.HabilitaMenues(False)
    Call frmMain.HabilitaBotones(False)
    DoEvents
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
                
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
        
        On Local Error GoTo Salir
        
        'cargar tablas
        lvwProcs.ListItems.Clear
        
        Busqueda = Trim$(txtBuscar.Text)
        
        If Len(Trim$(Busqueda)) = 0 Or Busqueda = "*.*" Then
            'cargar todas
            For K = 1 To UBound(cState(ActiveConexion).aProcs)
                Set itmx = lvwProcs.ListItems.Add(, , cState(ActiveConexion).aProcs(K).Nombre, 5, 5)
                itmx.SubItems(1) = cState(ActiveConexion).aProcs(K).Descripción
            Next K
        Else
            j = 1
            For K = 1 To UBound(cState(ActiveConexion).aProcs)
                If cState(ActiveConexion).aProcs(K).Nombre Like Busqueda Then
                    Set itmx = lvwProcs.ListItems.Add(, , cState(ActiveConexion).aProcs(K).Nombre, 5, 5)
                    itmx.SubItems(1) = cState(ActiveConexion).aProcs(K).Descripción
                    j = j + 1
                End If
            Next K
            
            If lvwProcs.ListItems.Count = 0 Then
                MsgBox "No se han encontrado procedimientos con esa coincidencia.", vbCritical
            End If
        End If
        
        Call Hourglass(hwnd, False)
    End If
    
    Set ci = Nothing
    Set itmx = Nothing
    
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    
    Exit Sub:
Salir:
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Function CaracterEspecial(ByVal Caracter As String) As Boolean

    Dim ret As Boolean
    Dim K As Integer
    
    ret = False
    
    For K = 1 To UBound(aSentencias)
        If aSentencias(K).Glosa = Caracter Then
            ret = True
            Exit For
        End If
    Next K

    CaracterEspecial = ret
    
End Function

Public Sub DeleteProc()

    Dim ActiveConexion As Integer
    Dim K As Integer
    Dim Campo As Field
    Dim Biblio As String
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwProcs.SelectedItem Is Nothing Then
            
            ActiveConexion = ConexionActiva(imgConexiones.SelectedItem.Text)
            
            Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaProcedures)
                
            Do While Not Rs.EOF
                If Rs!procedure_name = lvwProcs.SelectedItem.Text Then
                    For Each Campo In Rs.Fields
                        If Campo.Name = "PROCEDURE_SCHEMA" Then
                            If Not IsNull(Campo.Value) Then
                                Biblio = Campo.Value
                            End If
                        End If
                    Next
                    Exit Do
                End If
                Rs.MoveNext
            Loop
                        
            Rs.Close
            
            If Biblio <> "" Then
                txtQuery.Text = "DROP PROCEDURE " & Biblio & "." & lvwProcs.SelectedItem.Text
            Else
                txtQuery.Text = "DROP PROCEDURE " & lvwProcs.SelectedItem.Text
            End If
            Call FormateaSentencias
            
            frmMain.EjecutaQuery
                    
        Else
            MsgBox "Debe seleccionar un procedimiento.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
    Set Campo = Nothing
    
End Sub

Public Sub EjecutaProc()

    On Local Error GoTo ErrorEjecutaProc
    
    Dim ActiveConexion As Integer
    Dim K As Integer
    Dim Biblio As String
    Dim Campo As Field
        
    Call Hourglass(hwnd, True)
    Call frmMain.HabilitaMenues(False)
    Call frmMain.HabilitaBotones(False)
            
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwProcs.SelectedItem Is Nothing Then
            ActiveConexion = ConexionActiva(imgConexiones.SelectedItem.Text)
            
            If cState(ActiveConexion).tipo = TIPO_ODBC Then
                Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaProcedures)
                    
                Do While Not Rs.EOF
                    DoEvents
                    If Rs!procedure_name = lvwProcs.SelectedItem.Text Then
                        For Each Campo In Rs.Fields
                            If Campo.Name = "PROCEDURE_SCHEMA" Then
                                If Not IsNull(Campo.Value) Then
                                    Biblio = Campo.Value
                                    Exit For
                                End If
                            End If
                        Next
                            
                        Rs.Close
                                            
                        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaProcedureParameters)
                                            
                        K = 0
                        Do While Not Rs.EOF
                            If Rs!procedure_name = lvwProcs.SelectedItem.Text Then
                                K = K + 1
                                Exit Do
                            End If
                            Rs.MoveNext
                        Loop
                        
                        If K = 0 Then
                            If Biblio <> "" Then
                                frmMain.ActiveForm.txtQuery.Text = "CALL " & Biblio & "." & lvwProcs.SelectedItem.Text
                            Else
                                frmMain.ActiveForm.txtQuery.Text = "CALL " & lvwProcs.SelectedItem.Text
                            End If
                            frmMain.EjecutaQuery
                        Else
                            frmRunProc.CONEXION_ORIGEN = imgConexiones.SelectedItem.Text
                            frmRunProc.TABLA_ORIGEN = lvwProcs.SelectedItem.Text
                            frmRunProc.Show vbModal
                        End If
                        Exit Do
                    End If
                    Rs.MoveNext
                Loop
                            
                Rs.Close
            Else
                'rescatar info desde .mdb
                Set cat.ActiveConnection = DBConnection(ActiveConexion)
                Set cmdx = cat.Procedures(lvwProcs.SelectedItem.Text).Command
                
                'ver cuantos parametros hay
                K = 0
                For Each prm In cmdx.Parameters
                    K = K + 1
                    Exit For
                Next
                
                'si hay un parametro ejecutar inmediatamente
                If K = 0 Then
                    frmMain.ActiveForm.txtQuery.Text = cmdx.CommandText
                    frmMain.ActiveForm.FormateaSentencias
                    frmMain.EjecutaQuery
                Else
                    frmRunProc.CONEXION_ORIGEN = imgConexiones.SelectedItem.Text
                    frmRunProc.TABLA_ORIGEN = lvwProcs.SelectedItem.Text
                    frmRunProc.Show vbModal
                End If
                
                Set cat = Nothing
                Set cmdx = Nothing
                Set prm = Nothing
                
            End If
        Else
            MsgBox "Debe seleccionar un procedimiento.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Call Hourglass(hwnd, False)
    
    Exit Sub
    
ErrorEjecutaProc:
    MsgBox "EjecutaProc : " & Err & " " & Error$, vbCritical
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Err = 0
    
End Sub

Public Sub EliminarTabla()
       
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwTablas.SelectedItem Is Nothing Then
        
            If InStr(1, lvwTablas.SelectedItem.Text, " ") = 0 Then
                txtQuery.Text = "DROP TABLE " & lvwTablas.SelectedItem.Text
            Else
                txtQuery.Text = "DROP TABLE [" & lvwTablas.SelectedItem.Text & "]"
            End If
            
            Call FormateaSentencias
            
            If frmMain.EjecutaQuery() Then
                lvwTablas.ListItems.Remove lvwTablas.SelectedItem.Key
                Call ActualizarTablas
            End If
        Else
            MsgBox "Debe seleccionar una tabla.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub


Public Sub EliminarVista()
       
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwVistas.SelectedItem Is Nothing Then
        
            If InStr(1, lvwVistas.SelectedItem.Text, " ") = 0 Then
                txtQuery.Text = "DROP VIEW " & lvwVistas.SelectedItem.Text
            Else
                txtQuery.Text = "DROP VIEW [" & lvwVistas.SelectedItem.Text & "]"
            End If
            
            Call FormateaSentencias
            
            If frmMain.EjecutaQuery() Then
                lvwVistas.ListItems.Remove lvwVistas.SelectedItem.Key
                Call ActualizarVistas
            End If
        Else
            MsgBox "Debe seleccionar una vista.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 And Not Shift Then
        Call frmMain.EjecutaQuery
    ElseIf KeyCode = 32 Then
        Call FormateaSentencias
    End If
            
End Sub


Private Sub Form_Load()

    On Local Error Resume Next
    
    Line (0, 0)-(ScaleWidth, 0), vb3DShadow
    Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
        
    nContador = 0
    
    HayCambios = True
    
    txtQuery.Font.Name = glbFontLetra
    txtQuery.Font.Size = glbFontSize
    txtQuery.SelColor = glbColorSql
    txtQuery.BackColor = glbBackColor
    
    Call FormateaSentencias
            
    Call CargaHistorial(Me.lvwHistorial)
    
    Err = 0
    
End Sub

Public Sub FormateaSentencias()

    On Local Error Resume Next
    
    Dim sql As String
    Dim ret As Long
    Dim K As Integer
    Dim pos As Integer
    Dim ntokens As Integer
    
    Dim CHARRANGE As ECharrange
        
    sql = Replace(txtQuery.Text, Chr$(11), "")
    sql = Replace(txtQuery.Text, Chr$(13), "")
    sql = Replace(txtQuery.Text, Chr$(0), "")
    
    pos = txtQuery.SelStart
    
    If sql = "" Then
        Exit Sub
    End If
    
    txtQuery.Visible = False
    
    ret = SendMessage(txtQuery.hwnd, EM_EXGETSEL, 0, CHARRANGE)
    
    txtQuery.SelStart = 1
    txtQuery.SelFontName = glbFontLetra
    txtQuery.SelLength = Len(txtQuery.Text)
    txtQuery.SelColor = glbColorSql
    txtQuery.SetFocus
    txtQuery.SelBold = False
    txtQuery.SelBold = False
        
    For K = 1 To UBound(aSentencias)
        Call ColorSQL(txtQuery, aSentencias(K).Glosa, aSentencias(K).Color)
    Next K
    
    HayCambios = False
        
    txtQuery.Visible = True
    txtQuery.SelStart = pos
    txtQuery.SelLength = 0
    txtQuery.SetFocus
    txtQuery.SelColor = glbColorSql
    txtQuery.SelBold = False
    
    Err = 0
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not glbSalir Then
        Call frmMain.EliminarHojaConsulta
    End If
    
End Sub

Private Sub Form_Resize()
    
    On Local Error Resume Next
    
    Dim nWidth As Long
    Dim nWidth2 As Long
    Dim nHeight As Long
    Dim nHeight2 As Long
    Dim nTop As Long
    
    If Me.WindowState <> vbMinimized Then
    
        If WindowState = vbMaximized Then Call LockWindowUpdate(hwnd)
        DoEvents
        
        nHeight = Me.Height - staQuery.Height - 500
        nWidth = Me.Width - (Me.Width - (Splitter.Left + Splitter.Width)) - 180
        nTop = 1815
        
        Splitter.Height = nHeight
        
        'panel izquierdo
        lblOrigenes.Move 90, 90, nWidth
        imgConexiones.Move 90, 330, nWidth
        lblInfoCon.Move 90, 705, nWidth
        tabInfo.Move 90, 1005, nWidth, nHeight - 930
        fraBus.Move 135, 1380, nWidth - 100 ', nHeight - 1300
        txtBuscar.Move 810, 45, fraBus.Width - 900
        lviewCampos.Move 135, nTop, nWidth - 100, nHeight - 1800
        lvwTablas.Move 135, nTop, nWidth - 100, nHeight - 1800
        lvwVistas.Move 135, nTop, nWidth - 100, nHeight - 1800
        lvwProcs.Move 135, nTop, nWidth - 100, nHeight - 1800
        'lvwIndices.Move 135, nTop, nWidth - 100, nHeight - 1800
        
        griQuery.ReDraw = False
        
        nWidth2 = Me.ScaleWidth - nWidth - 180
        TabQuery.Move Splitter.Left + 100, 90, nWidth2, nHeight
        
        nHeight2 = TabQuery.Height - 500
        
        txtQuery.Move Splitter.Left + 150, 510, nWidth2 - 100, nHeight2
        griQuery.Move Splitter.Left + 150, 510, nWidth2 - 100, nHeight2
        lviewDetalle.Move Splitter.Left + 150, 510, nWidth2 - 100, nHeight2
        lvwHistorial.Move Splitter.Left + 150, 510, nWidth2 - 100, nHeight2 / 2
        txtHist.Move Splitter.Left + 150, (lvwHistorial.Top + lvwHistorial.Height) + 20, nWidth2 - 100, (nHeight2 / 2) '+ 10
                        
        griQuery.Refresh
        griQuery.ReDraw = True
        
        If Splitter.Left < 1000 Then
            Splitter.Left = 3750
            Call Form_Resize
        End If
        
        If (Splitter.Left > Me.Width + 100) Or ((Me.Width - Splitter.Left) <= 2500) Then
            Splitter.Left = 4000
            Call Form_Resize
        End If
        
        pgbQuery.Left = staQuery.Panels(1).Left + 1
        pgbQuery.Width = staQuery.Panels(1).Width - 1
        pgbQuery.Top = Me.ScaleHeight - 260
        pgbQuery.Move staQuery.Panels(1).Left + 1, Me.ScaleHeight - 260, staQuery.Panels(1).Width - 1
                
        Splitter.ZOrder 0
        If WindowState = vbMaximized Then Call LockWindowUpdate(0&)
    End If
    
    Err = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Local Error Resume Next
        
    fState(Val(Mid$(Me.Tag, 2))).Deleted = True
    
    Call AnimaWindow(hwnd, 50)
    
    DoEvents
    
    frmMain.Arrange vbTileHorizontal
    
    Set frmQuery = Nothing
    
    Err = 0
    
End Sub



Private Sub griQuery_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim Sel As Variant
    Dim nsel As Long
    
    Call griQuery.GetText(1, Row, Sel)
    
    If Sel = "1" Then
        griQuery.Row = Row
        griQuery.Col = 2
        griQuery.Row2 = Row
        griQuery.Col2 = griQuery.MaxCols
        griQuery.BackColor = QBColor(14)
        griQuery.BlockMode = True
    Else
        griQuery.Row = Row
        griQuery.Col = 2
        griQuery.Row2 = Row
        griQuery.Col2 = griQuery.MaxCols
        griQuery.BackColor = griQuery.SelForeColor
        griQuery.BlockMode = True
    End If
    
End Sub

Private Sub griQuery_Click(ByVal Col As Long, ByVal Row As Long)

    Dim K As Integer
    Dim itmx As ListItem
    Dim Valor As Variant
    Dim Columna As String
    
    If lviewCampos.ListItems.Count > 0 And Row > 0 Then
        lviewCampos.ListItems.Clear
        lviewCampos.Sorted = False
        'llenar los campos
        For K = 2 To griQuery.MaxCols
            Call griQuery.GetText(K, 0, Valor)
            
            Set itmx = frmMain.ActiveForm.lviewCampos.ListItems.Add()
            itmx.Text = Valor
            itmx.Icon = 3
            itmx.SmallIcon = 3
        Next K
        
        'llenar los datos en la grilla
        For K = 2 To griQuery.MaxCols
            Call griQuery.GetText(K, Row, Valor)
            Call griQuery.GetText(K, Row, Valor)
            Set itmx = lviewCampos.ListItems(K - 1)
            itmx.SubItems(1) = Valor
        Next K
        tabInfo.Tabs(1).Selected = True
    End If
    
    Set itmx = Nothing
    
End Sub

Private Sub griQuery_DblClick(ByVal Col As Long, ByVal Row As Long)

    If griQuery.DataRowCnt > 0 Then
        lviewCampos.ListItems(Col - 1).Selected = True
        Call lviewCampos_DblClick
    End If
    
End Sub

Private Sub griQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If griQuery.DataRowCnt > 0 Then
            If frmMain.mnuQuery.Enabled Then
                PopupMenu frmMain.mnuQuery
            End If
        End If
    End If
    
End Sub


Private Sub imgConexiones_Click()

    Call CargaInfoProveedor
    
End Sub

Public Sub CargaInfoProveedor()

    Dim ci As ComboItem
    Dim itmx As ListItem
    Dim ActiveConexion As Integer
    Dim K  As Integer
    
    Call frmMain.HabilitaMenues(False)
    Call frmMain.HabilitaBotones(False)
        
    If Not imgConexiones.SelectedItem Is Nothing Then
        Call Hourglass(hwnd, True)
        
        Set ci = imgConexiones.SelectedItem
        ActiveConexion = ConexionActiva(ci.Text)
        
        On Local Error GoTo Salir
        
        'cargar tablas
        lvwTablas.ListItems.Clear
        For K = 1 To UBound(cState(ActiveConexion).aTablas)
            Set itmx = lvwTablas.ListItems.Add(, "t" & K, cState(ActiveConexion).aTablas(K).Nombre, 5, 5)
            itmx.SubItems(1) = cState(ActiveConexion).aTablas(K).Descripción
        Next K
        
        tabInfo.Tabs(2).Caption = "&Tablas (" & lvwTablas.ListItems.Count & ")"
        
        'cargar vistas
        lvwVistas.ListItems.Clear
        For K = 1 To UBound(cState(ActiveConexion).aVistas)
            Set itmx = lvwVistas.ListItems.Add(, "v" & K, cState(ActiveConexion).aVistas(K).Nombre, 8, 8)
            itmx.SubItems(1) = cState(ActiveConexion).aVistas(K).Descripción
        Next K
        
        tabInfo.Tabs(3).Caption = "&Vistas (" & lvwVistas.ListItems.Count & ")"
        
        'cargar procs
        lvwProcs.ListItems.Clear
        For K = 1 To UBound(cState(ActiveConexion).aProcs)
            Set itmx = lvwProcs.ListItems.Add(, "p" & K, cState(ActiveConexion).aProcs(K).Nombre, 7, 7)
            itmx.SubItems(1) = cState(ActiveConexion).aProcs(K).Descripción
        Next K
                
        tabInfo.Tabs(4).Caption = "&Procedimientos (" & lvwProcs.ListItems.Count & ")"
        
        'cargar indices
        'lvwIndices.ListItems.Clear
        'For k = 1 To UBound(cState(ActiveConexion).aIndice)
        '    Set itmx = lvwIndices.ListItems.Add(, "i" & k, cState(ActiveConexion).aIndice(k).Nombre, 7, 7)
            'Itmx.SubItems(1) = cState(ActiveConexion).aIndice(K).Descripción
        'Next k
        
        'tabInfo.Tabs(5).Caption = "&Indices (" & lvwIndices.ListItems.Count & ")"
                
        Call Hourglass(hwnd, False)
    End If
    
    Set ci = Nothing
    Set itmx = Nothing
    
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    
    Exit Sub:
Salir:
    Call frmMain.HabilitaMenues(True)
    Call frmMain.HabilitaBotones(True)
    Call frmMain.HabiBotones2
    Err = 0

End Sub

Private Sub imgConexiones_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Private Sub lviewCampos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lviewCampos.SortOrder = lvwAscending Then
        lviewCampos.SortOrder = lvwDescending
    Else
        lviewCampos.SortOrder = lvwAscending
    End If
    
    lviewCampos.Sorted = True
    
End Sub

Private Sub lviewCampos_DblClick()
    
    If Not lviewCampos.SelectedItem Is Nothing Then
        If lviewCampos.SelectedItem.SubItems(1) <> "" Then
            frmVisor.Show vbModal
        End If
    End If
    
End Sub

Private Sub lviewDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lviewDetalle.SortOrder = lvwAscending Then
        lviewDetalle.SortOrder = lvwDescending
    Else
        lviewDetalle.SortOrder = lvwAscending
    End If
    
    lviewDetalle.Sorted = True
    
End Sub

Private Sub lviewDetalle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If lviewDetalle.ListItems.Count > 0 Then
            PopupMenu frmMain.mnuTablas
        End If
    End If
    
End Sub


Private Sub lvwHistorial_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwHistorial.SortOrder = lvwAscending Then
        lvwHistorial.SortOrder = lvwDescending
    Else
        lvwHistorial.SortOrder = lvwAscending
    End If
    
    lvwHistorial.Sorted = True

End Sub

Private Sub lvwHistorial_DblClick()
    
    On Local Error Resume Next
    
    If Not lvwHistorial.SelectedItem Is Nothing Then
        txtQuery.Text = lvwHistorial.ListItems(lvwHistorial.SelectedItem.Key).SubItems(1)
        'HayCambios = True
        Call FormateaSentencias
        TabQuery.Tabs(1).Selected = True
    End If
    
    Err = 0
    
End Sub

Private Sub lvwHistorial_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Local Error Resume Next
    
    If Not lvwHistorial.SelectedItem Is Nothing Then
        txtHist.Text = lvwHistorial.ListItems(lvwHistorial.SelectedItem.Key).SubItems(1)
    End If
    
    Err = 0
    
End Sub


Private Sub lvwProcs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwProcs.SortOrder = lvwAscending Then
        lvwProcs.SortOrder = lvwDescending
    Else
        lvwProcs.SortOrder = lvwAscending
    End If
    
    lvwProcs.Sorted = True
    
End Sub

Private Sub lvwProcs_DblClick()

    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwProcs.SelectedItem Is Nothing Then
            Call EjecutaProc
        Else
            MsgBox "Debe seleccionar un procedimiento.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub lvwProcs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu frmMain.mnuProcs
    End If
    
End Sub


Private Sub lvwTablas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwTablas.SortOrder = lvwAscending Then
        lvwTablas.SortOrder = lvwDescending
    Else
        lvwTablas.SortOrder = lvwAscending
    End If
    
    lvwTablas.Sorted = True
    
End Sub

Private Sub lvwTablas_DblClick()

    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwTablas.SelectedItem Is Nothing Then
            If InStr(lvwTablas.SelectedItem.Text, " ") = 0 Then
                txtQuery.Text = "SELECT * FROM " & lvwTablas.SelectedItem.Text
            Else
                txtQuery.Text = "SELECT * FROM [" & lvwTablas.SelectedItem.Text & "]"
            End If
            Call FormateaSentencias
            frmMain.EjecutaQuery
        Else
            MsgBox "Debe seleccionar una tabla.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub lvwTablas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu frmMain.mnuQTabla
    End If
    
End Sub


Private Sub lvwVistas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwVistas.SortOrder = lvwAscending Then
        lvwVistas.SortOrder = lvwDescending
    Else
        lvwVistas.SortOrder = lvwAscending
    End If
    
    lvwVistas.Sorted = True
    
End Sub

Private Sub lvwVistas_DblClick()

    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwVistas.SelectedItem Is Nothing Then
            If InStr(lvwVistas.SelectedItem.Text, " ") = 0 Then
                txtQuery.Text = "SELECT * FROM " & lvwVistas.SelectedItem.Text
            Else
                txtQuery.Text = "SELECT * FROM [" & lvwVistas.SelectedItem.Text & "]"
            End If
            'HayCambios = True
            Call FormateaSentencias
            frmMain.EjecutaQuery
        Else
            MsgBox "Debe seleccionar una vista.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Public Sub InfoDeCampos()
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwTablas.SelectedItem Is Nothing Then
            Dim ci As ComboItem
        
            Set ci = imgConexiones.SelectedItem
            
            Call frmMain.HabilitaMenues(False)
            Call frmMain.HabilitaBotones(False)
            DoEvents
            
            frmVerTabla.TABLA_ORIGEN = lvwTablas.SelectedItem.Text
            frmVerTabla.CONEXION_ORIGEN = ci.Text
            frmVerTabla.Show vbModal
            
            Call frmMain.HabilitaMenues(True)
            Call frmMain.HabilitaBotones(True)
            Call frmMain.HabiBotones2
    
            Set ci = Nothing
        Else
            MsgBox "Debe seleccionar una tabla.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Public Sub InfoDeCamposVista()
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwVistas.SelectedItem Is Nothing Then
            Dim ci As ComboItem
        
            Set ci = imgConexiones.SelectedItem
            
            Call frmMain.HabilitaMenues(False)
            Call frmMain.HabilitaBotones(False)
                
            frmVerTabla.VISTA = lvwVistas.SelectedItem.Text
            frmVerTabla.CONEXION_ORIGEN = ci.Text
            
            frmVerTabla.Show vbModal
            
            Call frmMain.HabilitaMenues(True)
            Call frmMain.HabilitaBotones(True)
            Call frmMain.HabiBotones2
    
            Set ci = Nothing
        Else
            MsgBox "Debe seleccionar una vista.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub


Public Sub QuickView()
    
    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwTablas.SelectedItem Is Nothing Then
            lvwTablas_DblClick
        Else
            MsgBox "Debe seleccionar una tabla.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub
Public Sub QuickViewVista()

    If Not imgConexiones.SelectedItem Is Nothing Then
        If Not lvwVistas.SelectedItem Is Nothing Then
            lvwVistas_DblClick
        Else
            MsgBox "Debe seleccionar una vista.", vbCritical
        End If
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub
Private Sub lvwVistas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu frmMain.mnuVista
    End If
    
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' if the left button is down set the flag
    If Button = 1 Then fInitiateDrag = True
End Sub


Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the flag isn't set then the left button wasn't
    ' pressed while the mouse was over one of the splitters
    If fInitiateDrag <> True Then Exit Sub

    ' if the left button is down then we want to move the splitter
    If Button = 1 Then ' if the Tag is false then we need to set
        If Splitter.Tag = False Then ' the color and clip the cursor.
    
            Splitter.BackColor = &H808080 '<- set the "dragging" color here
          
            Splitter.Tag = True
        End If
    
        Splitter.Move (Splitter.Left + x) - (SPLT_WDTH \ 3)
    End If
    
    Splitter.Refresh
    
End Sub


Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the left button is the one being released we need to reset
    ' the color, Tag, flag, cancel ClipCursor and call form_resize
  
    If Button = 1 Then           ' to move the list and text boxes
        Splitter.Tag = False
        fInitiateDrag = False
        'ClipCursor ByVal 0&
        Splitter.BackColor = &H8000000F  '<- set to original color
        Form_Resize
    End If
    
End Sub


Private Sub tabInfo_Click()

    If tabInfo.SelectedItem.Index = 1 Then
        lviewCampos.ZOrder 0
    ElseIf tabInfo.SelectedItem.Index = 2 Then
        lvwTablas.ZOrder 0
        lvwTablas.Visible = True
    ElseIf tabInfo.SelectedItem.Index = 3 Then
        lvwVistas.ZOrder 0
        lvwVistas.Visible = True
    ElseIf tabInfo.SelectedItem.Index = 4 Then
        lvwProcs.ZOrder 0
        lvwProcs.Visible = True
    End If
    
    Splitter.ZOrder 0
    
End Sub

Private Sub TabQuery_Click()
    
    If TabQuery.SelectedItem.Index = 1 Then
        txtQuery.ZOrder 0
        txtQuery.Visible = True
        griQuery.Visible = False
        lviewDetalle.Visible = False
        lvwHistorial.Visible = False
        txtHist.Visible = False
    ElseIf TabQuery.SelectedItem.Index = 2 Then
        griQuery.ZOrder 0
        griQuery.Visible = True
        txtQuery.Visible = False
        lviewDetalle.Visible = False
        lvwHistorial.Visible = False
        txtHist.Visible = False
    ElseIf TabQuery.SelectedItem.Index = 3 Then
        lviewDetalle.ZOrder 0
        lviewDetalle.Visible = True
        griQuery.Visible = False
        txtQuery.Visible = False
        lvwHistorial.Visible = False
        txtHist.Visible = False
    Else
        lvwHistorial.ZOrder 0
        lvwHistorial.Visible = True
        txtHist.Visible = True
        txtHist.ZOrder 0
        lviewDetalle.Visible = False
        griQuery.Visible = False
        txtQuery.Visible = False
    End If
    
End Sub







Private Sub txtBuscar_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If tabInfo.SelectedItem.Index = 2 Then
            Call BuscarTabla
        ElseIf tabInfo.SelectedItem.Index = 3 Then
            Call BuscarVista
        ElseIf tabInfo.SelectedItem.Index = 4 Then
            Call BuscarProcedimiento
        End If
    End If
    
End Sub

Private Sub txtQuery_KeyPress(KeyAscii As Integer)

    Dim ret As Integer
    Dim ret2 As Long
    Dim CHARRANGE As ECharrange
        
    If KeyAscii <> vbKeyReturn Then
        If Len(txtQuery.SelText) > 0 Then
            If KeyAscii <> 3 Then
                If txtQuery.Visible Then
                    txtQuery.SelText = ""
                End If
            End If
        End If
        
        If CaracterEspecial(Chr$(KeyAscii)) Then
            flag = True
        End If
    
        ret = SendMessage(txtQuery.hwnd, EM_EXGETSEL, 0, CHARRANGE)
        txtQuery.SelStart = CHARRANGE.cpMin
        txtQuery.SelLength = 0
        txtQuery.SelFontName = glbFontLetra
        txtQuery.SelFontSize = glbFontSize
        txtQuery.SelColor = glbColorSql
        txtQuery.SelBold = False
    Else
        Call FormateaSentencias
    End If
        
End Sub


Private Sub txtQuery_KeyUp(KeyCode As Integer, Shift As Integer)
    If flag Then
        flag = False
        Call FormateaSentencias
    End If
End Sub

Private Sub txtQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu frmMain.mnuEdicion
    End If
    
End Sub



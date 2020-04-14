VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Easy Query !"
   ClientHeight    =   6120
   ClientLeft      =   1920
   ClientTop       =   4005
   ClientWidth     =   9705
   Icon            =   "EasyQuery.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgVB6"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   33
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdConectar"
            Object.ToolTipText     =   "Conectar a origen de datos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdDesconectar"
            Object.ToolTipText     =   "Desconectar origen de datos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdIniTrx"
            Object.ToolTipText     =   "Iniciar transacción"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdRollTrx"
            Object.ToolTipText     =   "Deshacer transacción"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdFinTrx"
            Object.ToolTipText     =   "Finalizar transacción"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdNuevo"
            Object.ToolTipText     =   "Nuevo query"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEliminar"
            Object.ToolTipText     =   "Eliminar query"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdLimpiar"
            Object.ToolTipText     =   "Limpiar query"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAscendente"
            Object.ToolTipText     =   "Ordenar columnas ascendentemente"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdDescendente"
            Object.ToolTipText     =   "Ordenar columnas descendentemente"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAbrir"
            Object.ToolTipText     =   "Abrir archivo de texto"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdGuardar"
            Object.ToolTipText     =   "Guardar texto a archivo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdImprimir"
            Object.ToolTipText     =   "Imprimir texto a impresora"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCopiar"
            Object.ToolTipText     =   "Copia texto al portapapeles"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCortar"
            Object.ToolTipText     =   "Corta el texto al portapapeles"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPegar"
            Object.ToolTipText     =   "Pega el texto del portapapeles"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdDeshacer"
            Object.ToolTipText     =   "Deshace los cambios hechos"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdBuscar"
            Object.ToolTipText     =   "Ejecuta la sentencia sql"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdQStop"
            Object.ToolTipText     =   "Detener ejecución de sql"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdDetener"
            Object.ToolTipText     =   "Parar carga de registros"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCascade"
            Object.ToolTipText     =   "Organizar ventanas en cascada"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTileH"
            Object.ToolTipText     =   "Organizar ventanas horizontalmente"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTileV"
            Object.ToolTipText     =   "Organizar ventanas verticalmente"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Salir de la aplicación"
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgVB6 
      Left            =   615
      Top             =   1650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":105E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":11BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":1316
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":142A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":1D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2022
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2136
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":224A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2472
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2586
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":269A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":27AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":28C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":29D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":2E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":33B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":3512
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":366E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EasyQuery.frx":37CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5850
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivox 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuQuery_Conectar 
         Caption         =   "&Conectar origen de datos"
      End
      Begin VB.Menu mnuQuery_Desconectar 
         Caption         =   "&Desconectar origen de datos"
      End
      Begin VB.Menu mnuQuery_Sep99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Nuevahoja 
         Caption         =   "&Nuevo Query"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuQuery_EliminarHoja 
         Caption         =   "&Eliminar Query"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuQuery_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_AbrirconsultaSQL 
         Caption         =   "&Abrir consulta SQL"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuQuery_GuardarSQL 
         Caption         =   "&Guardar SQL"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuQuery_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_ConfigImpresora 
         Caption         =   "&Configurar Página"
      End
      Begin VB.Menu mnuQuery_ImprimirSQL 
         Caption         =   "&Imprimir SQL"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuQuery_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEdicion_Deshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEdicion_Copiar 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnuEdicion_Pegar 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu mnuEdicion_Cortar 
         Caption         =   "C&ortar"
      End
      Begin VB.Menu mnuEdicion_Borrar 
         Caption         =   "&Borrar"
      End
      Begin VB.Menu mnuEdicion_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_SeleccionarTodo 
         Caption         =   "&Seleccionar todo ..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEdicion_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Buscar 
         Caption         =   "B&uscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdicion_BuscarSiguiente 
         Caption         =   "Buscar Si&guiente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdicion_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Reemplazar 
         Caption         =   "&Reemplazar"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
      Enabled         =   0   'False
      Begin VB.Menu mnuQuery_SeleccionarTodo 
         Caption         =   "&Seleccionar todos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuQuery_DesSeleccionarTodo 
         Caption         =   "&Cancelar selección"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuQuery_Invertir 
         Caption         =   "&Invertir selección"
      End
      Begin VB.Menu mnuQuery_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_CopiarDatos 
         Caption         =   "&Copiar fila ..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuQuery_EiminarFila 
         Caption         =   "&Eliminar filas ..."
      End
      Begin VB.Menu mnuQuery_Imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuQuery_CopiarHeader 
         Caption         =   "Copiar &Header"
      End
      Begin VB.Menu sex1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_VerTabla 
         Caption         =   "&Ver estructura de tabla"
      End
      Begin VB.Menu mnuQuery_VerTodasTablas 
         Caption         =   "Ver todas las &tablas de conexión"
      End
   End
   Begin VB.Menu mnuExportar 
      Caption         =   "&Exportar"
      Enabled         =   0   'False
      Begin VB.Menu mnuQuery_ExportarCSV 
         Caption         =   "Exportar a .CSV"
      End
      Begin VB.Menu mnuQuery_ExportarXML 
         Caption         =   "Exportar a .XML"
      End
      Begin VB.Menu mnuQuery_ExportarTXT 
         Caption         =   "Exportar a .TXT"
      End
      Begin VB.Menu mnuQuery_ExportarHTM 
         Caption         =   "Exportar a .HTM"
      End
      Begin VB.Menu mnuQuery_ExportarXLS 
         Caption         =   "Exportar a .XLS"
      End
      Begin VB.Menu mnuQuery_ExportarRTF 
         Caption         =   "Exportar a .RTF"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuery_ExportarTAB 
         Caption         =   "Exportar a .TAB"
      End
   End
   Begin VB.Menu mnuQTabla 
      Caption         =   "&Tablas"
      Enabled         =   0   'False
      Begin VB.Menu mnuQTabla_Agregar 
         Caption         =   "&Agregar Tabla"
      End
      Begin VB.Menu mnuQTabla_Qview 
         Caption         =   "&Consultar Tabla"
      End
      Begin VB.Menu mnuQTabla_InfoCampo 
         Caption         =   "Campos de &Tabla"
      End
      Begin VB.Menu mnuQTabla_Eliminar 
         Caption         =   "&Eliminar Tabla"
      End
      Begin VB.Menu mnuQTabla_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQTabla_Actualizar 
         Caption         =   "Actuali&zar"
      End
   End
   Begin VB.Menu mnuVista 
      Caption         =   "&Vistas"
      Enabled         =   0   'False
      Begin VB.Menu mnuVista_Agregar 
         Caption         =   "&Agregar Vista"
      End
      Begin VB.Menu mnuVista_Consultar 
         Caption         =   "&Consultar Vista"
      End
      Begin VB.Menu mnuVista_Campos 
         Caption         =   "Campos de &Vista"
      End
      Begin VB.Menu mnuVista_Eliminar 
         Caption         =   "&Eliminar Vista"
      End
      Begin VB.Menu mnuVista_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVista_Actualizar 
         Caption         =   "Actuali&zar"
      End
   End
   Begin VB.Menu mnuProcs 
      Caption         =   "&Procedimientos"
      Enabled         =   0   'False
      Begin VB.Menu mnuProcs_Add 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu mnuProcs_Editar 
         Caption         =   "E&ditar"
      End
      Begin VB.Menu mnuProcs_Del 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuProcs_Run 
         Caption         =   "E&jecutar"
      End
      Begin VB.Menu mnuProcs_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProcs_Actualizar 
         Caption         =   "Actuali&zar"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_Configuracion 
         Caption         =   "&Configurar Comandos críticos"
      End
      Begin VB.Menu mnuOpciones_ConfQuery 
         Caption         =   "Configurar &Editor de Querys"
      End
      Begin VB.Menu mnuOpciones_Historial 
         Caption         =   "Configurar &Historial de Querys"
      End
      Begin VB.Menu mnuOpciones_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Colores 
         Caption         =   "Colores sentencias &SQL"
      End
      Begin VB.Menu mnuOpciones_Skin 
         Caption         =   "&Skin"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpciones_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Browser 
         Caption         =   "&Browser de Querys"
      End
      Begin VB.Menu mnuOpciones_FTexto 
         Caption         =   "&Formato de Texto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpciones_SiempreVisible 
         Caption         =   "Siempre &visible"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuVentana_Cascada 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuVentana_Horizontal 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu mnuVentana_Vertical 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu mnuVentana_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentana_Organizar 
         Caption         =   "&Organizar ventanas"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAyuda_Contenido 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "&Indice"
      End
      Begin VB.Menu mnuAyuda_Busqueda 
         Caption         =   "&Búsqueda"
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_Web 
         Caption         =   "&Easy Query ! en el WEB"
      End
      Begin VB.Menu mnuAyuda_AcercaDe 
         Caption         =   "&Acerca de Easy Query ! ..."
      End
   End
   Begin VB.Menu mnuTablas 
      Caption         =   "Tablas"
      Visible         =   0   'False
      Begin VB.Menu mnuTablas_ImprimirCampos 
         Caption         =   "Imprimir campos consulta"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents xConnection As ADODB.Connection
Attribute xConnection.VB_VarHelpID = -1
Private WithEvents xRecordset As ADODB.Recordset
Attribute xRecordset.VB_VarHelpID = -1

Public m_Detener As Boolean
Public m_Cancelar As Boolean
Public m_Ejecutando As Boolean

Private itmx As ListItem



Private Sub AbreSql()

    On Local Error GoTo ErrorAbreSql
    
    Dim cc As New GCommonDialog
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim sql As String
    Dim Glosa As String
    
    If Not frmMain.ActiveForm Is Nothing Then
        Glosa = "Archivos de comandos Sql (*.SQL)|*.SQL|"
        Glosa = Glosa & "Archivos de Texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        
        If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , App.Path, "Abrir archivo ...", "SQL", Me.hwnd)) Then
           Exit Sub
        End If
                
        If Archivo <> "" Then
            nArchivo = FreeFile
            
            Open Archivo For Input As #nArchivo
                sql = Input(LOF(nArchivo), nArchivo)
            Close #nArchivo
            frmMain.ActiveForm.txtQuery.Text = vbNullString
            frmMain.ActiveForm.txtQuery.SelColor = glbColorSql
            frmMain.ActiveForm.txtQuery.Text = sql
        End If
        
        'frmMain.ActiveForm.txtQuery.Visible = False
        frmMain.ActiveForm.HayCambios = True
                    
        Call frmMain.ActiveForm.FormateaSentencias
        
        'frmMain.ActiveForm.txtQuery.Visible = True
        frmMain.ActiveForm.txtQuery.SetFocus
    End If
    
    Set cc = Nothing
    
    GoTo SalirAbreSql
    
ErrorAbreSql:
    Resume SalirAbreSql
    
SalirAbreSql:
    Err = 0
    
End Sub



Private Sub Ascendente()

    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.griQuery.DataRowCnt > 0 Then
                    
            frmMain.ActiveForm.griQuery.Col = 1 'frmMain.ActiveForm.griQuery.ActiveCol
            frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
            frmMain.ActiveForm.griQuery.Row = 0
            frmMain.ActiveForm.griQuery.Row2 = frmMain.ActiveForm.griQuery.DataRowCnt
            
            frmMain.ActiveForm.griQuery.SortBy = 0
            frmMain.ActiveForm.griQuery.SortKey(1) = frmMain.ActiveForm.griQuery.ActiveCol
            frmMain.ActiveForm.griQuery.SortKeyOrder(1) = 1
            frmMain.ActiveForm.griQuery.Action = 25
            
        End If
    End If
    
End Sub

Private Sub CargarInformacionDeCampos(ByRef NCampos As Integer)

    Dim Nombre As String
    Dim Campo As Field
        
    NCampos = 2
    
    With xRecordset
        'configurar celdas de la spread segun tipos de campos del recordset
        Call ConfiguraCeldas(xRecordset, NCampos)
                    
        'llenar listview de campos
        For Each Campo In .Fields
            Nombre = Campo.Name
            Set itmx = frmMain.ActiveForm.lviewDetalle.ListItems.Add()
            
            itmx.Text = Nombre
            itmx.Icon = 3
            itmx.SmallIcon = 3
            
            itmx.SubItems(1) = TipoDeCampo(Campo.Type)
            itmx.SubItems(2) = Campo.DefinedSize
            itmx.SubItems(3) = Campo.Precision
            
            If (Campo.Attributes And adFldIsNullable) Then
                itmx.SubItems(4) = "Si"
            Else
                itmx.SubItems(4) = "No"
            End If
        Next
    End With
    
    Set Campo = Nothing
    
End Sub
'cargar registros
Private Sub CargarRegistros(ByRef NCampos As Integer)

    Dim e As Long
    Dim f As Long
    Dim K As Integer
    
    f = 1
    
    'ciclar x los registros devueltos
    With xRecordset
        Do While Not .EOF
            e = DoEvents()
            
            'detener carga de registros
            If m_Detener Then
                m_Detener = False
                Exit Do
            End If
            
            'cada 100 filas incrementar
            If f Mod 100 = 0 Then
                frmMain.ActiveForm.griQuery.MaxRows = frmMain.ActiveForm.griQuery.MaxRows + 100
            End If
            
            'llenar x columnas
            For K = 2 To NCampos - 1
                If Not IsNull(.Fields(K - 2)) Then
                    Call frmMain.ActiveForm.griQuery.SetText(K, f, .Fields(K - 2))
                Else
                    Call frmMain.ActiveForm.griQuery.SetText(K, f, "NULO")
                End If
            Next K
                                
            frmStop.lblRegistros.Caption = CStr(f)
            
            f = f + 1
            .MoveNext
        Loop
        frmMain.ActiveForm.griQuery.MaxRows = frmMain.ActiveForm.griQuery.DataRowCnt
        frmMain.ActiveForm.griQuery.ReDraw = True
    End With
                
End Sub

Private Function ChequeaComandoCritico() As Boolean

    Dim K As Integer
    Dim ret As Boolean
    
    Dim sql As String
    
    sql = Trim$(ClearEnterInString(frmMain.ActiveForm.txtQuery.Text))
    
    ret = False
    
    For K = 1 To UBound(Comandos)
        If Comandos(K).Activo Then
            If UCase$(Trim$(Comandos(K).Comando)) = UCase$(Left$(sql, Len(Comandos(K).Comando))) Then
                ret = True
                Exit For
            End If
        End If
    Next K
    
    ChequeaComandoCritico = ret
    
End Function

'configura la celda de la spread
Private Sub ConfiguraCeldas(Rs As ADODB.Recordset, ByRef NCampos As Integer)

    Dim Campo As Field
    Dim Nombre As String
    Dim sCampo As String
    Dim Flagx As Boolean
    
    'cargar headers de la spread
    frmMain.ActiveForm.griQuery.BlockMode = False
    
    For Each Campo In Rs.Fields
        If NCampos > frmMain.ActiveForm.griQuery.MaxCols Then
            frmMain.ActiveForm.griQuery.MaxCols = frmMain.ActiveForm.griQuery.MaxCols + 1
        End If
        Nombre = Campo.Name
        Set itmx = frmMain.ActiveForm.lviewCampos.ListItems.Add()
        itmx.Text = Nombre
        itmx.Icon = 3
        itmx.SmallIcon = 3
        
        frmMain.ActiveForm.griQuery.Col = NCampos
        frmMain.ActiveForm.griQuery.Row = 0
        frmMain.ActiveForm.griQuery.Text = Nombre
                
        frmMain.ActiveForm.griQuery.Col = NCampos
        
        sCampo = TipoDeCampo(Campo.Type)
        Flagx = False
        
        Select Case sCampo
            Case "BigInt", "Currency", "Decimal", "Double", "Integer", "LongVarBinary", _
                 "Numeric", "Single", "SmallInt", "TinyInt", "UnsignedBigInt", _
                "UnsignedInt", "UnsignedSmallInt", "UnsignedTinyInt"
            
                Flagx = True
                frmMain.ActiveForm.griQuery.Row = -1
                frmMain.ActiveForm.griQuery.Row2 = -1
                frmMain.ActiveForm.griQuery.Col = NCampos
                frmMain.ActiveForm.griQuery.Col2 = NCampos
                frmMain.ActiveForm.griQuery.CellType = CellTypeStaticText
                frmMain.ActiveForm.griQuery.TypeHAlign = 1
                frmMain.ActiveForm.griQuery.TypeVAlign = 2
        End Select
        
        NCampos = NCampos + 1
    Next
                
End Sub

Private Sub Descendente()

    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.griQuery.DataRowCnt > 0 Then
                    
            frmMain.ActiveForm.griQuery.Col = 1 'frmMain.ActiveForm.griQuery.ActiveCol
            frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
            frmMain.ActiveForm.griQuery.Row = 0
            frmMain.ActiveForm.griQuery.Row2 = frmMain.ActiveForm.griQuery.DataRowCnt
            
            frmMain.ActiveForm.griQuery.SortBy = 0
            frmMain.ActiveForm.griQuery.SortKey(1) = frmMain.ActiveForm.griQuery.ActiveCol
            frmMain.ActiveForm.griQuery.SortKeyOrder(1) = 2
            frmMain.ActiveForm.griQuery.Action = 25
            
        End If
    End If
    
End Sub

Private Sub DetenerSql()

    On Local Error Resume Next
    
    Call Hourglass(hwnd, True)
    
    m_Cancelar = True
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Cancelando ejecución de SQL"
    xRecordset.Cancel
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecución de sentencia SQL cancelada con éxito."
    
    Err = 0
    
End Sub

Public Sub EjecutaQuery2(Optional ByVal cmdx As ADODB.Command)

    On Local Error GoTo ErrorEjecutaQuery
    
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim Msg As String
    Dim e As Long
    
    'verificar cual es el origen de datos activo
    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    ActiveConexion = ConexionActiva(ci.Text)
        
    'limpiar info
    Call LimpiarGrilla
        
    'requiere confirmacion del usuario ?
    If ChequeaComandoCritico() Then
        Msg = "Confirma ejecutar este comando crítico definido."
        If Confirma(Msg) = vbNo Then GoTo Salir
    End If
    
    'comenzar
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)
    
    Toolbar.Buttons("cmdQStop").Enabled = True
    Toolbar.Buttons("cmdSalir").Enabled = False
    frmMain.ActiveForm.txtQuery.Enabled = False
        
    'encender flag de ejecucion
    m_Ejecutando = True
                
    'preparar conexion y recordset
    Set xConnection = New ADODB.Connection
    Set xConnection = DBConnection(ActiveConexion)
    xConnection.CommandTimeout = glbTimeOut
    
    Set xRecordset = New ADODB.Recordset
    Set xRecordset.ActiveConnection = xConnection
    xRecordset.CursorLocation = adUseServer
    
    If cState(ActiveConexion).tipo = TIPO_ODBC Then
        Set xRecordset = xConnection.Execute(frmMain.ActiveForm.txtQuery.Text)
    Else
        Set xRecordset = cmdx.Execute
    End If
    
    'ejecutar sql
    'esperar hasta que conteste el equipo
    On Local Error Resume Next
    
    Do While xRecordset.State = adStateExecuting
        e = DoEvents
        If m_Cancelar Then
            Exit Do
        End If
    Loop
    Err = 0
        
    'carga de campos y de registros
    Call SQLStart
                            
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Listo."
        
    GoTo Salir
    
ErrorEjecutaQuery:
    If Not frmMain.ActiveForm Is Nothing Then
        If Err > 0 And Not m_Cancelar Then
            frmMain.ActiveForm.staQuery.Panels(2).Text = Error$
        End If
    End If
    
    Resume Salir
    
Salir:
    Unload frmStop
        
    Err = 0
    
    Call Hourglass(hwnd, False)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call HabilitaMenues(True)
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.txtQuery.Enabled = True
        On Local Error Resume Next
        frmMain.ActiveForm.txtQuery.SetFocus
        Err = 0
    End If
            
    m_Ejecutando = False
    m_Detener = False
    m_Cancelar = False
    
    'liberar memoria
    If xRecordset.State > 0 Then xRecordset.Close
    Set xRecordset = Nothing
    Set xConnection = Nothing
    Set ci = Nothing
    
End Sub

Public Sub EliminarHojaConsulta()

    On Local Error Resume Next
            
    Dim fIndex
    
    fIndex = frmMain.ActiveForm.Tag
    
    Unload frmMain.ActiveForm
    
    If Not IsNumeric(fIndex) Then
        If mnuOpciones_Browser.Checked Then
            frmBrowser.treQuerys.Nodes.Remove fIndex
        End If
    Else
        If mnuOpciones_Browser.Checked Then
            frmBrowser.treQuerys.Nodes.Remove "q1"
        End If
    End If
    Err = 0
            
End Sub

Private Sub FinTRX()
    
    On Local Error Resume Next
    
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim Msg As String
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then Exit Sub

    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
            
    ActiveConexion = ConexionActiva(ci.Text)
    
    If cState(ActiveConexion).Trx = True Then
        
        Msg = "Confirma finalizar transacción activa."
        If Confirma(Msg) = vbNo Then Exit Sub
        
        cState(ActiveConexion).Trx = False
        DBConnection(ActiveConexion).CommitTrans
        
        Toolbar.Buttons("cmdIniTrx").Enabled = True
        Toolbar.Buttons("cmdFinTrx").Enabled = False
        Toolbar.Buttons("cmdRollTrx").Enabled = False
    End If
    
    Set ci = Nothing
    Err = 0
            
End Sub

Public Sub GrabaQuerys()

    Dim K As Integer
    Dim qR As Integer
    Dim sql As String
    Dim ArrayCount As Integer
    Dim i As Integer
    Dim f As Form
    Dim NComandos
    'Dim sComando
    Dim Activo As String
    
    qR = 1
    
    Call Hourglass(hwnd, True)
    
    If UBound(Document) = 0 Then Exit Sub
    
    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = ArrayCount To 1 Step -1
        If Not fState(i).Deleted Then
            Set f = Document(i)
            sql = f.txtQuery.Text
            If sql <> "" Then
                Call f.txtQuery.SaveFile(App.Path & "\query" & qR & ".sql", 1)
                qR = qR + 1
            End If
            Unload f
        End If
    Next
    
    Call GrabaIni(C_INI, "Querys", "nquerys", CStr(qR - 1))
    
    ReDim Document(0)
    
    Call GrabaIni(C_INI, "ArchivosMDB", "nArchivosMdb", UBound(aArchivosMdb))
        
    For K = 1 To UBound(aArchivosMdb)
        Call GrabaIni(C_INI, "ArchivosMDB", "Archivo" & K, aArchivosMdb(K))
    Next K
    
    'grabar comandos
    
    'leer comandos criticos
    NComandos = UBound(Comandos)
         
    If NComandos > 0 Then
        For K = 1 To NComandos
            If Comandos(K).Activo Then
                Activo = "1"
            Else
                Activo = "0"
            End If
            
            Call GrabaIni(C_INI, "Comandos", "comando" & K, Comandos(K).Comando & "," & Activo)
        Next K
    End If
    
    Call GrabaIni(C_INI, "Comandos", "ncomandos", NComandos)
    
    Call Hourglass(hwnd, False)
    
End Sub




Private Sub GrabaSQL()

    On Local Error GoTo ErrorGrabaSQL
    
    Dim cc As New GCommonDialog
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim sql As String
    Dim Glosa As String
        
    If Not frmMain.ActiveForm Is Nothing Then
        Glosa = "Archivos de comandos Sql (*.SQL)|*.SQL|"
        Glosa = Glosa & "Archivos de comandos Texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        
        If Not (cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar como ...", "SQL", Me.hwnd)) Then
           Exit Sub
        End If
               
        If Archivo <> "" Then
            nArchivo = FreeFile
            sql = frmMain.ActiveForm.txtQuery.Text
            Archivo = StripNulls(Archivo)
            Open Archivo For Output As #nArchivo
                Print #nArchivo, sql
            Close #nArchivo
        End If
    End If
    
    Set cc = Nothing
    
    GoTo SalirGrabaSQL
    
ErrorGrabaSQL:
    Resume SalirGrabaSQL
    
SalirGrabaSQL:
    Err = 0
    
End Sub


Public Sub HabiBotones2()

    Toolbar.Buttons("cmdIniTrx").Enabled = Toolbar.Buttons("cmdIniTrx").Tag
    Toolbar.Buttons("cmdFinTrx").Enabled = Toolbar.Buttons("cmdFinTrx").Tag
    Toolbar.Buttons("cmdRollTrx").Enabled = Toolbar.Buttons("cmdRollTrx").Tag
    Toolbar.Buttons("cmdQStop").Enabled = False
    Toolbar.Buttons("cmdDetener").Enabled = False
    Toolbar.Buttons("cmdSalir").Enabled = True
    
End Sub

Public Sub HabilitaBotones(ByVal Estado As Boolean)

    Dim K As Integer
    
    If Not Estado Then
        Toolbar.Buttons("cmdIniTrx").Tag = Toolbar.Buttons("cmdIniTrx").Enabled
        Toolbar.Buttons("cmdFinTrx").Tag = Toolbar.Buttons("cmdFinTrx").Enabled
        Toolbar.Buttons("cmdRollTrx").Tag = Toolbar.Buttons("cmdRollTrx").Enabled
    End If
    
    For K = 1 To Toolbar.Buttons.Count
        Toolbar.Buttons(K).Enabled = Estado
    Next K
    
End Sub

Public Sub HabilitaMenues(ByVal Estado As Boolean)

    Dim K As Integer
    
    For K = 0 To Me.Controls.Count - 1
        If TypeOf Me.Controls(K) Is Menu Then
            If Me.Controls(K).Caption <> "-" Then
                Me.Controls(K).Enabled = Estado
            End If
        End If
    Next K
    
End Sub


Private Sub Imprimir()

    On Local Error GoTo ErrorImprimir
    
    If frmMain.ActiveForm.txtQuery.Text = "" Then Exit Sub
        
    Call Hourglass(hwnd, True)
    
    If Not ShowPrinter(Me) Then GoTo SalirImprimir
    
    Printer.Print frmMain.ActiveForm.txtQuery.Text
        
    GoTo SalirImprimir
    
ErrorImprimir:
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub IniciarTRX()

    On Local Error Resume Next
            
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then Exit Sub

    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    
    ActiveConexion = ConexionActiva(ci.Text)
    
    If cState(ActiveConexion).Trx = False Then
        cState(ActiveConexion).Trx = True
        DBConnection(ActiveConexion).BeginTrans
        Toolbar.Buttons("cmdIniTrx").Enabled = False
        Toolbar.Buttons("cmdFinTrx").Enabled = True
        Toolbar.Buttons("cmdRollTrx").Enabled = True
    End If
    
    Set ci = Nothing
    Err = 0
    
End Sub

Public Sub InstalaSistemaAyuda()

    'If glbPathSkin = "" Then
    '    Set HelpObj = New HelpCallBack
    
    '    Call Install(Me.hwnd, , imgList, , True)
      
    '    Call FontName(Me.hwnd, "Verdana")
    '    Call FontSize(Me.hwnd, 8&)
    'End If
    
End Sub

Private Sub LeeQuerys()

    Dim K As Integer
    Dim sql As String
    Dim c As Integer
    Dim sComando
    Dim Activo
    Dim Sentencia As String
    Dim Color
    Dim Nquerys
    Dim NComandos
    Dim Nsentencias
    Dim nArchivosMdb
    Dim nFreeFile As Long
        
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)
    
VolverS:
    Nsentencias = LeeIni("Sentencias", "nSentencias", C_INI)
    If Nsentencias = "" Then
        Call GrabaIni(C_INI, "Sentencias", "Sentencia1", "SELECT")
        Call GrabaIni(C_INI, "Sentencias", "Color1", "255")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia2", "FROM")
        Call GrabaIni(C_INI, "Sentencias", "Color2", "16744448")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia3", "WHERE")
        Call GrabaIni(C_INI, "Sentencias", "Color3", "16744703")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia4", "GROUP BY")
        Call GrabaIni(C_INI, "Sentencias", "Color4", "12615680")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia5", "ORDER BY")
        Call GrabaIni(C_INI, "Sentencias", "Color5", "12615808")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia6", "'")
        Call GrabaIni(C_INI, "Sentencias", "Color6", "4227072")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia7", ",")
        Call GrabaIni(C_INI, "Sentencias", "Color7", "8454143")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia8", ".")
        Call GrabaIni(C_INI, "Sentencias", "Color8", "8454016")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia9", "AND")
        Call GrabaIni(C_INI, "Sentencias", "Color9", "8421440")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia10", "=")
        Call GrabaIni(C_INI, "Sentencias", "Color10", "16711808")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia11", "(")
        Call GrabaIni(C_INI, "Sentencias", "Color11", "8421631")
        Call GrabaIni(C_INI, "Sentencias", "Sentencia12", ")")
        Call GrabaIni(C_INI, "Sentencias", "Color12", "8421631")
        Call GrabaIni(C_INI, "Sentencias", "nSentencias", "12")
        Nsentencias = 12
    End If
    
    ReDim aSentencias(Val(Nsentencias))
    
    For K = 1 To Nsentencias
        Sentencia = LeeIni("Sentencias", "Sentencia" & K, C_INI)
        Color = LeeIni("Sentencias", "Color" & K, C_INI)
        aSentencias(K).Glosa = Sentencia
        aSentencias(K).Color = CLng(Color)
    Next K
      
    glbTimeOut = LeeIni("General", "timeout", C_INI)
    If glbTimeOut = "" Then glbTimeOut = "0"
    
    Nquerys = LeeIni("Querys", "nquerys", C_INI)
    If Nquerys = "" Then Nquerys = 0
     
    If Nquerys > 0 Then
        For K = 1 To Nquerys
        
            On Error Resume Next
            
            nFreeFile = FreeFile
            sql = vbNullString
            Open App.Path & "\query" & K & ".sql" For Input As #nFreeFile
                sql = Input$(LOF(nFreeFile), nFreeFile)
            Close #nFreeFile
                        
            If sql <> "" Then
                If K > 1 Then
                    Call FileNew
                Else
                    fState(1).Tag = "q1"
                    fState(1).Dirty = 1
                End If
                Document(K).txtQuery.Text = sql
                Document(K).HayCambios = True
                Document(K).FormateaSentencias
                Document(K).Tag = "q" & K
            End If
            
            Err = 0
            
        Next K
    End If
    
    nArchivosMdb = LeeIni("ArchivosMDB", "nArchivosMdb", C_INI)
    If nArchivosMdb = "" Then nArchivosMdb = 0
    
    ReDim aArchivosMdb(Val(nArchivosMdb))
    
    For K = 1 To nArchivosMdb
        aArchivosMdb(K) = LeeIni("ArchivosMDB", "Archivo" & K, C_INI)
    Next K
    
VolverC:
    'leer comandos criticos
    NComandos = LeeIni("Comandos", "ncomandos", C_INI)
    If NComandos = "" Then NComandos = 0
     
    If NComandos > 0 Then
        c = 1
        For K = 1 To NComandos
            sComando = LeeIni("Comandos", "comando" & K, C_INI)
            
            If sComando <> "" Then
                ReDim Preserve Comandos(c)
                
                Comandos(c).Comando = Left$(sComando, Len(sComando) - 2)
                If Right$(sComando, 1) = "0" Then
                    Comandos(c).Activo = False
                Else
                    Comandos(c).Activo = True
                End If
                c = c + 1
            End If
        Next K
    Else
        Call GrabaIni(C_INI, "Comandos", "ncomandos", "4")
        Call GrabaIni(C_INI, "Comandos", "comando1", "DELETE,1")
        Call GrabaIni(C_INI, "Comandos", "comando2", "UPDATE,1")
        Call GrabaIni(C_INI, "Comandos", "comando3", "ALTER TABLE,1")
        Call GrabaIni(C_INI, "Comandos", "comando4", "DROP TABLE,1")
        GoTo VolverC
    End If
    
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
    
End Sub


Private Sub Limpiar()

    On Local Error Resume Next
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.griQuery.MaxRows = 1
        frmMain.ActiveForm.griQuery.MaxCols = 1
        
        Call frmMain.ActiveForm.griQuery.SetText(1, 0, "")
        Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
        Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
        frmMain.ActiveForm.txtQuery.Text = ""
        frmMain.ActiveForm.txtQuery.SetFocus
        Err = 0
    End If
    
End Sub

Private Sub LimpiarGrilla()

    frmMain.ActiveForm.griQuery.MaxRows = 1
    frmMain.ActiveForm.griQuery.MaxCols = 1
    
    Call frmMain.ActiveForm.griQuery.SetText(1, 0, "Selec")
    Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
    Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
    Call frmMain.ActiveForm.lviewDetalle.ListItems.Clear
    
End Sub

Private Sub PegarTexto()

    On Local Error Resume Next
            
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm!txtQuery.SelText = Clipboard.GetText(rtfText)
        frmMain.ActiveForm.HayCambios = True
        Call frmMain.ActiveForm.FormateaSentencias
        Err = 0
    End If
    
End Sub

Private Sub RollTRX()
    
    On Local Error Resume Next
    
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim Msg As String
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then Exit Sub

    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem

    ActiveConexion = ConexionActiva(ci.Text)
    
    If cState(ActiveConexion).Trx = True Then
        
        Msg = "Confirma cancelar transacción activa."
        If Confirma(Msg) = vbNo Then Exit Sub
        
        cState(ActiveConexion).Trx = False
        DBConnection(ActiveConexion).RollbackTrans
        
        Toolbar.Buttons("cmdIniTrx").Enabled = True
        Toolbar.Buttons("cmdFinTrx").Enabled = False
        Toolbar.Buttons("cmdRollTrx").Enabled = False
    End If
    Err = 0
            
    Set ci = Nothing
    
End Sub

'ejecutar sql
Private Sub SqlRun()

    Dim e As Long
    
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando Query ...."
    
    xRecordset.Open glbSQl, xConnection, adOpenForwardOnly, adLockReadOnly, adCmdText + adAsyncExecute
                                
    'esperar hasta que conteste el equipo
    On Local Error Resume Next
    
    Do While xRecordset.State = adStateExecuting
        e = DoEvents
        If m_Cancelar Then
            Exit Do
        End If
    Loop
    Err = 0
    
End Sub

Private Sub SQLStart()

    Dim NCampos As Integer
    
    'trajo registros la consulta ?
    If Not m_Cancelar Then
        If xRecordset.State > 0 Then
            If Not xRecordset.EOF Then
                'apurar ejecucion con 100 filas
                frmMain.ActiveForm.griQuery.MaxRows = 100
                frmMain.ActiveForm.TabQuery.Tabs(2).Selected = True
                
                'llenar campos con listview
                Call CargarInformacionDeCampos(NCampos)
                
                'cargar datos en la spread
                frmMain.ActiveForm.griQuery.ReDraw = False
                Toolbar.Buttons("cmdDetener").Enabled = True
                    
                Toolbar.Buttons("cmdQStop").Enabled = False
                
                'pantalla de detener
                frmStop.Show
                DoEvents
                    
                'cargar registros
                Call CargarRegistros(NCampos)
            End If
            
            frmMain.ActiveForm.griQuery.Row = 1
            
            'cerrar y liberar
            
            frmMain.ActiveForm.staQuery.Panels(2).Text = CStr(frmMain.ActiveForm.griQuery.DataRowCnt) & " registros cargados."
        Else
            frmMain.ActiveForm.staQuery.Panels(2).Text = "Sentencia SQL no devolvio registros."
        End If
    End If
        
End Sub


Private Sub Undo()

    On Local Error Resume Next
    
    Dim CanUndo As Boolean
    
    If Not frmMain.ActiveForm Is Nothing Then
        CanUndo = SendMessageByVal(frmMain.ActiveForm.txtQuery.hwnd, EM_CANUNDO, 0, 0)
        
        If CanUndo Then
            SendMessageByVal frmMain.ActiveForm.txtQuery.hwnd, EM_UNDO, 0, 0
            'Call Deshacer(frmMain.ActiveForm.txtQuery.hwnd)
            frmMain.ActiveForm.HayCambios = True
            'frmMain.ActiveForm.txtQuery.Visible = False
            Call frmMain.ActiveForm.FormateaSentencias
            'frmMain.ActiveForm.txtQuery.Visible = True
        End If
        Err = 0
    End If
    
End Sub

'validar datos a exportar
Private Function ValidaExportar() As Boolean

    Dim ret As Boolean
    Dim Sel As Variant
    Dim K As Long
    
    ret = False
    
    If frmMain.ActiveForm.imgConexiones.ComboItems.Count > 0 Then
        If Not frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then
            If frmMain.ActiveForm.griQuery.DataRowCnt > 0 Then
                For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
                    Call frmMain.ActiveForm.griQuery.GetText(1, K, Sel)
                    If Sel = "1" Then
                        ret = True
                        Exit For
                    End If
                Next K
                
                If ret = False Then
                    MsgBox "Debe seleccionar registros a exportar.", vbCritical
                End If
            Else
                MsgBox "No hay registros a exportar.", vbCritical
            End If
        Else
            MsgBox "Debe seleccionar un origen de datos.", vbCritical
        End If
    Else
        MsgBox "Debe conectarse a un origen de datos.", vbCritical
    End If
    
    ValidaExportar = ret
    
End Function

'validar antes de ejecutar
Private Function Validar() As Boolean

    Dim ret As Boolean
    
    ret = False
    
    'se esta conectado
    If Not Conectado Then
        MsgBox "Debe conectarse a un origen de datos primero.", vbCritical
        GoTo Salir
    End If
    
    'hay form activo
    If frmMain.ActiveForm Is Nothing Then
        MsgBox "Debe activar un formulario o agregar una ventana de consultas.", vbCritical
        GoTo Salir
    End If
    
    'hay conexion seleccionada ?
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then
        MsgBox "Seleccione un origen de datos.", vbCritical
        frmMain.ActiveForm.imgConexiones.SetFocus
        GoTo Salir
    End If
    
    'verificar que se ejecuta
    glbSQl = ""
    
    If Len(frmMain.ActiveForm.txtQuery.SelText) = 0 Then
        glbSQl = Trim$(ClearEnterInString(Trim$(frmMain.ActiveForm.txtQuery.Text)))
    Else
        glbSQl = Trim$(ClearEnterInString(Trim$(frmMain.ActiveForm.txtQuery.SelText)))
    End If
    
    If glbSQl = "" Then
        MsgBox "Debe digitar una sentencia sql.", vbCritical
        frmMain.ActiveForm.txtQuery.SetFocus
        GoTo Salir
    End If
    
    ret = True
    
Salir:
    Validar = ret
    
End Function
Private Sub MDIForm_Load()
    
    Dim K As Integer
    Dim op
           
    Cargando = True
                
    On Local Error Resume Next
        
    Err = 0
    
    Call Hourglass(hwnd, True)
    
    Set xConnection = New ADODB.Connection
    Set xRecordset = New ADODB.Recordset
   
    gbLastPos = 0
    
    Toolbar.Buttons("cmdIniTrx").Tag = True
        
    ReDim Document(1)
    ReDim fState(1)
    
    ReDim DBConnection(0)
    ReDim cState(0)
    ReDim aArchivosMdb(0)
    ReDim Comandos(0)
    
    Document(1).Tag = "q1"
    
    fState(1).Dirty = False
    fState(K).Tag = "q1"
    fState(K).Dirty = 1
    
    Conectado = False
    
    RemoveMenus Me, False, False, _
        False, False, False, True, True
        
    Call LeeQuerys
    
    Cargando = False
    
    op = LeeIni("Opciones", "Browser", C_INI)
    
    If op = "-1" Then
        frmMain.mnuOpciones_Browser.Checked = False
        Call mnuOpciones_Browser_Click
    Else
        frmMain.mnuOpciones_Browser.Checked = True
        Call mnuOpciones_Browser_Click
    End If
        
    op = LeeIni("Opciones", "FTexto", C_INI)
    If op = "-1" Then
        frmMain.mnuOpciones_FTexto.Checked = False
        Call mnuOpciones_FTexto_Click
    Else
        frmMain.mnuOpciones_FTexto.Checked = True
        Call mnuOpciones_FTexto_Click
    End If
        
    Call Hourglass(hwnd, False)
    
End Sub
Public Function EjecutaQuery() As Boolean

    On Local Error GoTo ErrorEjecutaQuery
        
    Dim Msg As String
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim ret As Boolean
    
    ret = False
    
    'se esta ejecutando ya una sentencia
    If m_Ejecutando Then Exit Function
    
    'validar datos antes de ejecutar
    If Not Validar Then Exit Function
                
    'verificar cual es el origen de datos activo
    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    ActiveConexion = ConexionActiva(ci.Text)
        
    'limpiar info
    Call LimpiarGrilla
    
    'requiere confirmacion del usuario ?
    If ChequeaComandoCritico() Then
        Msg = "Confirma ejecutar este comando crítico definido."
        If Confirma(Msg) = vbNo Then GoTo Salir
    End If
    
    'comenzar
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)
    
    Toolbar.Buttons("cmdQStop").Enabled = True
    Toolbar.Buttons("cmdSalir").Enabled = False
    frmMain.ActiveForm.txtQuery.Enabled = False
        
    'encender flag de ejecucion
    m_Ejecutando = True
                
    'preparar conexion y recordset
    Set xConnection = New ADODB.Connection
    Set xConnection = DBConnection(ActiveConexion)
    xConnection.CommandTimeout = glbTimeOut
    
    Set xRecordset = New ADODB.Recordset
    Set xRecordset.ActiveConnection = xConnection
    xRecordset.CursorLocation = adUseServer
    
    'grabar historial
    Call GrabaHistorialSQL
    Call CargaHistorial(frmMain.ActiveForm.lvwHistorial)
                
    'si es un select ?
    If Left$(UCase$(glbSQl), 6) = "SELECT" Then
        'ejecutar sql
        Call SqlRun
        
        'carga de campos y de registros
        Call SQLStart
    Else
        'preparar para una ejecución ?
        Msg = "La sentencia sql escrita devuelve registros"
        
        If Confirma(Msg) = vbYes Then
            'ejecutar sql
            Call SqlRun
            
            'carga de campos y de registros
            Call SQLStart
        Else
            'ejecutar lo que el usuario decidio
            frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando instrucción SQL ...."
            Call xConnection.Execute(glbSQl)
            MsgBox "Acción ejecutada con éxito!", vbInformation
        End If
                        
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Listo."
    End If
    
    ret = True
    
    GoTo Salir
    
ErrorEjecutaQuery:
    If Not frmMain.ActiveForm Is Nothing Then
        If Err > 0 And Not m_Cancelar Then
            frmMain.ActiveForm.staQuery.Panels(2).Text = Error$
        End If
    End If
    ret = False
    Resume Salir
    
Salir:
    Unload frmStop
        
    Err = 0
    
    Call Hourglass(hwnd, False)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call HabilitaMenues(True)
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.txtQuery.Enabled = True
        On Local Error Resume Next
        frmMain.ActiveForm.txtQuery.SetFocus
        Err = 0
    End If
            
    m_Ejecutando = False
    m_Detener = False
    m_Cancelar = False
    
    'liberar memoria
    If xRecordset.State > 0 Then xRecordset.Close
    Set xRecordset = Nothing
    Set xConnection = Nothing
    Set ci = Nothing
            
    EjecutaQuery = ret
    
    
End Function
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim Msg As String
    
    If Not Toolbar.Buttons("cmdSalir").Enabled Then
        Cancel = 1
        Exit Sub
    End If
    
    Msg = "Confirma salir de " & App.Title
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
        
    DoEvents
    
    glbSalir = True
    
    Call GrabaQuerys
    
    Unload frmTablas
    Unload frmBrowser
    
    Set xRecordset = Nothing
    Set xConnection = Nothing
    
    DoEvents
    
    Me.Hide
            
End Sub
Private Sub MDIForm_Resize()

    If Me.WindowState = vbMinimized Then
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = False
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = False
    Else
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = True
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = True
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
        
    Unload frmTablas
    
    If mnuOpciones_Browser.Checked Then
        Call GrabaIni(C_INI, "Opciones", "Browser", "-1")
    Else
        Call GrabaIni(C_INI, "Opciones", "Browser", "0")
    End If
    
    If mnuOpciones_FTexto.Checked Then
        Call GrabaIni(C_INI, "Opciones", "FTexto", "-1")
    Else
        Call GrabaIni(C_INI, "Opciones", "FTexto", "0")
    End If
                
End Sub


Private Sub mnuArchivo_Acercade_Click()
    frmAcercaDe.Show vbModal
End Sub

Private Sub mnuArchivo_Restaurar_Click()

    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.Show
    
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuAyuda_AcercaDe_Click()
    frmAcercaDe.Show vbModal
End Sub


Private Sub mnuAyuda_Web_Click()

    On Local Error Resume Next
    ShellExecute Me.hwnd, vbNullString, "http://www.vbsoftware.cl/", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0

End Sub

Private Sub mnuEdicion_Borrar_Click()

    Dim Msg As String
    
    If Not frmMain.ActiveForm Is Nothing Then
        If Len(frmMain.ActiveForm.txtQuery.SelText) > 0 Then
            frmMain.ActiveForm.txtQuery.SelText = ""
        Else
            Msg = "Confirma borrar contenido"
            If Confirma(Msg) = vbYes Then
                frmMain.ActiveForm.txtQuery.Text = ""
            End If
        End If
    End If
    
End Sub

Private Sub mnuEdicion_Buscar_Click()
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmFind.Show
    End If
    
End Sub

Private Sub mnuEdicion_BuscarSiguiente_Click()
    Call FindText
End Sub


Private Sub mnuEdicion_Copiar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdCopiar"))
End Sub


Private Sub mnuEdicion_Cortar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdCortar"))
End Sub

Private Sub mnuEdicion_Deshacer_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdDeshacer"))
End Sub

Private Sub mnuEdicion_Pegar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdPegar"))
End Sub

Private Sub mnuEdicion_Reemplazar_Click()
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmReemplazar.Show
    End If
    
End Sub

Private Sub mnuEdicion_SeleccionarTodo_Click()
    
    On Local Error Resume Next
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.txtQuery.SelStart = 0
        frmMain.ActiveForm.txtQuery.SelLength = Len(frmMain.ActiveForm.txtQuery.Text)
        frmMain.ActiveForm.txtQuery.SetFocus
    End If
    
    Err = 0
    
End Sub

Private Sub mnuOpciones_Browser_Click()
    
    If Not Cargando Then
        mnuOpciones_Browser.Checked = Not mnuOpciones_Browser.Checked
        
        If mnuOpciones_Browser.Checked Then
            frmBrowser.Show
        Else
            Unload frmBrowser
        End If
    End If
    
End Sub



Private Sub mnuOpciones_Colores_Click()
     
     frmConfColores.Show vbModal
End Sub

Private Sub mnuOpciones_Configuracion_Click()
    frmOpciones.Show vbModal
End Sub

Private Sub mnuOpciones_ConfQuery_Click()
    frmConfQuery.Show vbModal
End Sub


Private Sub mnuOpciones_FTexto_Click()

    If Not Cargando Then
        mnuOpciones_FTexto.Checked = Not mnuOpciones_FTexto.Checked
        
        If mnuOpciones_FTexto.Checked Then
            'frmColores.Show
        Else
            'Unload frmColores
        End If
    End If
    
End Sub

Private Sub mnuOpciones_Historial_Click()
    frmHistorialQ.Show vbModal
End Sub

Private Sub mnuOpciones_SiempreVisible_Click()

    If Not Cargando Then
        mnuOpciones_SiempreVisible.Checked = Not mnuOpciones_SiempreVisible.Checked
        
        If mnuOpciones_SiempreVisible.Checked Then
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
        Else
            Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
        End If
    End If
    
End Sub

Private Sub mnuOpciones_Skin_Click()
   'frmConfSkin.Show vbModal
End Sub


Private Sub mnuProcs_Actualizar_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call frmMain.ActiveForm.ActualizarProcedimientos
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub mnuProcs_Add_Click()

    MsgBox "En desarrollo!", vbInformation
    
    Exit Sub
    
    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        frmMain.ActiveForm.AddProc
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
    
End Sub

Private Sub mnuProcs_Del_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        frmMain.ActiveForm.DeleteProc
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
    
End Sub

Private Sub mnuProcs_Editar_Click()
    MsgBox "En desarrollo!", vbInformation
End Sub

Private Sub mnuProcs_Run_Click()
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.EjecutaProc
    End If
End Sub

Private Sub mnuQTabla_Actualizar_Click()
    
    If Not frmMain.ActiveForm Is Nothing Then
        Call frmMain.ActiveForm.ActualizarTablas
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub mnuQTabla_Agregar_Click()
    MsgBox "En desarrollo!", vbInformation
End Sub

Private Sub mnuQTabla_Eliminar_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call frmMain.ActiveForm.EliminarTabla
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub mnuQTabla_InfoCampo_Click()
    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        Call frmMain.ActiveForm.InfoDeCampos
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
End Sub

Private Sub mnuQTabla_Qview_Click()
    
    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        Call frmMain.ActiveForm.QuickView
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
    
End Sub

Private Sub mnuQuery_AbrirconsultaSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdAbrir"))
End Sub

Private Sub mnuQuery_Conectar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdConectar"))
End Sub

Private Sub mnuQuery_ConfigImpresora_Click()
    
    Call ShowPageSetupDlg(hwnd)
        
End Sub

Private Sub mnuQuery_CopiarDatos_Click()

    On Local Error Resume Next
    
    Dim K As Integer
    Dim Sel
    
    Call Hourglass(hwnd, True)
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    frmMain.ActiveForm.griQuery.Col = 2
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row = 0
    frmMain.ActiveForm.griQuery.Row2 = 0
            
    gsBuffer = frmMain.ActiveForm.griQuery.Clip
    
    Clipboard.Clear
    
    For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.GetText(1, K, Sel)
        If Sel = "1" Then
            frmMain.ActiveForm.griQuery.Col = 2
            frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
            frmMain.ActiveForm.griQuery.Row = K
            frmMain.ActiveForm.griQuery.Row2 = K
            gsBuffer = gsBuffer & frmMain.ActiveForm.griQuery.Clip
        End If
    Next K
    
    Clipboard.Clear
    Clipboard.SetText gsBuffer, 1
    
    Call Hourglass(hwnd, False)
    
    Err = 0
    
End Sub

Private Sub mnuQuery_CopiarHeader_Click()

    Dim K As Integer
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    Clipboard.Clear
    
    frmMain.ActiveForm.griQuery.Row = 0
    
    For K = 2 To frmMain.ActiveForm.griQuery.MaxCols
        frmMain.ActiveForm.griQuery.Col = K
        gsBuffer = gsBuffer & frmMain.ActiveForm.griQuery.Text & " , "
    Next K
    
    If Len(gsBuffer) > 0 Then
        gsBuffer = Left$(gsBuffer, Len(gsBuffer) - 3)
    
        Clipboard.SetText gsBuffer, 1
    End If
    
End Sub

Private Sub mnuQuery_Desconectar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdDesconectar"))
End Sub

Private Sub mnuQuery_DesSeleccionarTodo_Click()

    On Local Error Resume Next
    
    Dim K As Long
    
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)
    
    If Not frmMain.ActiveForm Is Nothing Then
        Call Hourglass(hwnd, True)
            
        frmMain.ActiveForm.griQuery.ReDraw = False
        For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
            Call frmMain.ActiveForm.griQuery.SetText(1, K, "")
        Next K
        frmMain.ActiveForm.griQuery.ReDraw = True
        
        Call Hourglass(hwnd, False)
    End If
    
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
        
    Err = 0
    
End Sub

Private Sub mnuQuery_EiminarFila_Click()
        
    Dim K As Long
    Dim Sel
    Dim Elimine As Boolean
    
    Elimine = False
    If frmMain.ActiveForm.griQuery.ActiveRow > 0 Then
        For K = frmMain.ActiveForm.griQuery.MaxRows To 1 Step -1
            Call frmMain.ActiveForm.griQuery.GetText(1, K, Sel)
            If Sel = "1" Then
                frmMain.ActiveForm.griQuery.Row = K
                frmMain.ActiveForm.griQuery.Row2 = K
                frmMain.ActiveForm.griQuery.BlockMode = True
                frmMain.ActiveForm.griQuery.Action = ActionDeleteRow
                frmMain.ActiveForm.griQuery.BlockMode = False
                Elimine = True
            End If
        Next K
        
        If Not Elimine Then
            MsgBox "Debe seleccionar celdas a eliminar.", vbCritical
        Else
            frmMain.ActiveForm.griQuery.MaxRows = frmMain.ActiveForm.griQuery.DataRowCnt
            frmMain.ActiveForm.staQuery.Panels(2).Text = CStr(frmMain.ActiveForm.griQuery.DataRowCnt) & " registros."
        End If
    End If
    
End Sub

Private Sub mnuQuery_EliminarHoja_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdEliminar"))
End Sub

Private Sub mnuQuery_ExportarCSV_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de CSV (*.CSV)|*.CSV", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".csv"
            End If
            'exportar a csv
            If CSVExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarHTM_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de Hypertexto (*.HTM)|*.HTM", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".htm"
            End If
            'exportar a htm
            If HTMLExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarTAB_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de Tabulaciones (*.TAB)|*.TAB", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".tab"
            End If
            'exportar a tab
            If TABExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarTXT_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de texto (*.txt)|*.txt", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".txt"
            End If
            'exportar a txt
            If TXTExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarXLS_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de Excel (*.XLS)|*.XLS", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".xls"
            End If
            'exportar a htm
            If EXCELExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarXML_Click()

    Dim Archivo As String
    
    If ValidaExportar() Then
        Archivo = SaveDialog(hwnd, "Archivos de Sql (*.XML)|*.XML", "Guardar archivo como ...", App.Path)
        
        If StripNulls(Archivo) <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".xml"
            End If
            'exportar a csv
            If XMLExport(Archivo) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        End If
    End If
    
End Sub


Private Sub mnuQuery_GuardarSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdGuardar"))
End Sub

Private Sub mnuQuery_Imprimir_Click()
    
    Dim fChecked As Boolean
    
    fChecked = mnuOpciones_SiempreVisible.Checked
    
    If mnuOpciones_SiempreVisible.Checked Then
        Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    End If
        
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.ColHidden = True
    
    frmPreview.lHwnd = frmMain.ActiveForm.griQuery.hwnd
    frmPreview.Show vbModal
        
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.ColHidden = False
    
    If fChecked Then
        Call mnuOpciones_SiempreVisible_Click
    End If
    
End Sub

Private Sub mnuQuery_ImprimirSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdImprimir"))
End Sub

Private Sub mnuQuery_Invertir_Click()

    On Local Error Resume Next
    
    Dim K As Long
    Dim Sel As Variant
    
    If Not frmMain.ActiveForm Is Nothing Then
        Call Hourglass(hwnd, True)
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        
        frmMain.ActiveForm.griQuery.ReDraw = False
        For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
            Call frmMain.ActiveForm.griQuery.GetText(1, K, Sel)
            
            If Sel = "1" Then
                Call frmMain.ActiveForm.griQuery.SetText(1, K, "")
            Else
                Call frmMain.ActiveForm.griQuery.SetText(1, K, "1")
            End If
        Next K
        frmMain.ActiveForm.griQuery.ReDraw = True
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
        Call Hourglass(hwnd, False)
    End If
    
    Err = 0
    
End Sub

Private Sub mnuQuery_Nuevahoja_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdNuevo"))
End Sub


Private Sub mnuQuery_Salir_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdSalir"))
End Sub


Private Sub mnuQuery_SeleccionarTodo_Click()

    On Local Error Resume Next
    
    Dim K As Long
        
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)

    frmMain.ActiveForm.griQuery.ReDraw = False
    For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.SetText(1, K, "1")
    Next K
    frmMain.ActiveForm.griQuery.ReDraw = True
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    Clipboard.Clear
    
    frmMain.ActiveForm.griQuery.Col = 2
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row = -1
    frmMain.ActiveForm.griQuery.Row2 = -1
            
    gsBuffer = frmMain.ActiveForm.griQuery.Clip
    
    Clipboard.SetText gsBuffer, 1
    
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call Hourglass(hwnd, False)
    
    Err = 0
    
End Sub

Private Sub mnuQuery_VerTabla_Click()
      frmVerTabla.Show vbModal
End Sub

Private Sub mnuQuery_VerTodasTablas_Click()
    frmTablas.Show
End Sub

Private Sub mnuTablas_ImprimirCampos_Click()

    On Local Error GoTo ErrorImprimirCampos
    
    Dim itmx As ListItem
    Dim Header As String
    
    Dim K As Integer
    Dim j As Integer
    
    Dim Campos As Integer
    
    Dim buffer As String
    
    buffer = ""
    
    Campos = frmMain.ActiveForm.griQuery.MaxCols
        
    Printer.Font = "Courier New"
    Printer.Font.Size = 10
    
    Printer.FontBold = True
    Header = "Nombre Físico" & vbTab & "Tipo" & vbTab & "Largo" & vbTab & "Decimales" & vbTab & "Acepta Nulos"
    Printer.Print Header
    Printer.Print String$(60, "-")
    Printer.FontBold = False
    Printer.Print
    
    With frmMain.ActiveForm.lviewDetalle
        For K = 1 To .ListItems.Count
            Set itmx = .ListItems(K)
            
            buffer = itmx.Text & vbTab
            
            For j = 1 To 4
                If j = 3 Then
                    buffer = buffer & Space$(5 - Len(itmx.SubItems(j))) & itmx.SubItems(j) & vbTab
                ElseIf j = 4 Then
                    buffer = buffer & Space$(9 - Len(itmx.SubItems(j))) & itmx.SubItems(j) & vbTab
                ElseIf j = 5 Then
                    buffer = buffer & Space$(5 - Len(itmx.SubItems(j))) & itmx.SubItems(j)
                Else
                    buffer = buffer & itmx.SubItems(j) & vbTab
                End If
            Next j
            
            Printer.Print buffer
        Next K
    End With
    
    Printer.EndDoc
    
    GoTo SalirImprimirCampos
    
ErrorImprimirCampos:
    Printer.KillDoc
    MsgBox "ImprimirCampos : " & Err & " " & Error$, vbCritical
    Resume SalirImprimirCampos
    
SalirImprimirCampos:
    Err = 0
    
End Sub

Private Sub mnuVentana_Cascada_Click()
    
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)
    frmMain.Arrange vbCascade
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call Hourglass(hwnd, False)
End Sub

Private Sub mnuVentana_Horizontal_Click()
    
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)

    frmMain.Arrange vbTileHorizontal
    
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call Hourglass(hwnd, False)

End Sub


Private Sub mnuVentana_Organizar_Click()
    
    Call Hourglass(hwnd, True)
    Call HabilitaMenues(False)
    Call HabilitaBotones(False)

    frmMain.Arrange vbArrangeIcons
    
    Call HabilitaMenues(True)
    Call HabilitaBotones(True)
    Call HabiBotones2
    Call Hourglass(hwnd, False)

End Sub




Private Sub mnuVentana_Vertical_Click()
    frmMain.Arrange vbTileVertical
End Sub


Private Sub mnuVista_Actualizar_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call frmMain.ActiveForm.ActualizarVistas
    Else
        MsgBox "Debe seleccionar una conexión.", vbCritical
    End If
    
End Sub

Private Sub mnuVista_Agregar_Click()
    MsgBox "En desarrollo!", vbInformation
End Sub

Private Sub mnuVista_Campos_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        Call frmMain.ActiveForm.InfoDeCamposVista
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
    
End Sub

Private Sub mnuVista_Consultar_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call HabilitaMenues(False)
        Call HabilitaBotones(False)
        Call frmMain.ActiveForm.QuickViewVista
        Call HabilitaMenues(True)
        Call HabilitaBotones(True)
        Call HabiBotones2
    End If
    
End Sub

Private Sub mnuVista_Eliminar_Click()

    If Not frmMain.ActiveForm Is Nothing Then
        Call frmMain.ActiveForm.EliminarVista
    End If
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Dim fIndex
    Dim ActiveConexion As Integer
        
    Select Case Button.Key
        Case "cmdConectar"
            frmConexiones.Show vbModal
        Case "cmdDesconectar"
            frmConexiones.Show vbModal
        Case "cmdIniTrx"
            Call IniciarTRX
        Case "cmdFinTrx"
            Call FinTRX
        Case "cmdRollTrx"
            Call RollTRX
        Case "cmdNuevo"
            Call Hourglass(hwnd, True)
            Call HabilitaMenues(False)
            Call HabilitaBotones(False)
            
            Call FileNew(fIndex)
            
            If mnuOpciones_Browser.Checked = True Then
                frmBrowser.treQuerys.Nodes.Add("EJ1", tvwChild, _
                frmMain.ActiveForm.Tag, frmMain.ActiveForm.Caption, 1).EnsureVisible
            End If
            Call InsertaSentenciasPredefinidas
            mnuVentana_Horizontal_Click
            
            Call Hourglass(hwnd, False)
            Call HabilitaMenues(True)
            Call HabilitaBotones(True)
            Call HabiBotones2
        Case "cmdEliminar"
            Call Hourglass(hwnd, True)
            Call HabilitaMenues(False)
            Call HabilitaBotones(False)
            Call EliminarHojaConsulta
            Call Hourglass(hwnd, False)
            Call HabilitaMenues(True)
            Call HabilitaBotones(True)
            Call HabiBotones2
        Case "cmdLimpiar"
            Call Limpiar
        Case "cmdAscendente"
            Call Ascendente
        Case "cmdDescendente"
            Call Descendente
        Case "cmdAbrir"
            Call AbreSql
        Case "cmdGuardar"
            Call GrabaSQL
        Case "cmdImprimir"
            Call Imprimir
        Case "cmdCopiar"
            On Local Error Resume Next
            If Not frmMain.ActiveForm Is Nothing Then
                Call Copiar
            End If
            Err = 0
        Case "cmdCortar"
            On Local Error Resume Next
            If Not frmMain.ActiveForm Is Nothing Then
                Clipboard.SetText frmMain.ActiveForm!txtQuery.SelText
                frmMain.ActiveForm!txtQuery.SelText = ""
            End If
            Err = 0
        Case "cmdPegar"
            Call PegarTexto
        Case "cmdDeshacer"
            Call Undo
        Case "cmdBuscar"
            Call EjecutaQuery
        Case "cmdQStop"
            Call DetenerSql
        Case "cmdDetener"
            If Not frmMain.ActiveForm Is Nothing Then
                m_Detener = True
            End If
        Case "cmdCascade"
            mnuVentana_Cascada_Click
        Case "cmdTileH"
            mnuVentana_Horizontal_Click
        Case "cmdTileV"
            mnuVentana_Vertical_Click
        Case "cmdSalir"
            'Call GrabaQuerys
            Unload Me
    End Select
    
End Sub

Private Sub InsertaSentenciasPredefinidas()

    Dim sql As String
    
    sql = "SELECT * " & vbNewLine
    sql = sql & vbNewLine
    sql = sql & "FROM" & vbNewLine
    sql = sql & vbNewLine
    sql = sql & "WHERE" & vbNewLine
    sql = sql & vbNewLine
    sql = sql & "GROUP BY" & vbNewLine
    sql = sql & vbNewLine
    sql = sql & "ORDER BY" & vbNewLine
    sql = sql & vbNewLine
    
    frmMain.ActiveForm.txtQuery.Text = sql
    frmMain.ActiveForm.FormateaSentencias
    
    On Local Error Resume Next
    frmMain.ActiveForm.imgConexiones.ComboItems(1).Selected = True
    Err = 0
    
End Sub

Private Sub xConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
        
    Dim Msg As String
    
    If pError Is Nothing Then Exit Sub
        
    If pError.Number <> 0 Then
        
        If Not m_Cancelar Then
            Msg = pError.Source & vbNewLine & vbNewLine
            Msg = Msg & pError.Description & vbNewLine
            
            MsgBox Msg, vbCritical
        End If
        
    End If
    
End Sub

Private Sub xConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    
    Dim e As Integer
    
    Do While adStatus <> adStatusOK
        e = DoEvents()
    Loop
    
End Sub



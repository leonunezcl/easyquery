VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5EB28C2E-92F5-11D5-BD59-0000E82185CB}#1.0#0"; "XSYSTRAY.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Easy Query !"
   ClientHeight    =   6120
   ClientLeft      =   1800
   ClientTop       =   1920
   ClientWidth     =   8715
   Icon            =   "QueryIt.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   820
      ButtonWidth     =   688
      ButtonHeight    =   661
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   27
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdConectar"
            Object.ToolTipText     =   "Conecta hacia fuente de datos"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDesconectar"
            Object.ToolTipText     =   "Desconecta una conexión previa"
            Object.Tag             =   ""
            ImageIndex      =   20
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdIniTrx"
            Object.ToolTipText     =   "Iniciar Transacción"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdFinTrx"
            Object.ToolTipText     =   "Finalizar Transacción"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdRollTrx"
            Object.ToolTipText     =   "Cancelar Transaccion"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdNuevo"
            Object.ToolTipText     =   "Crea una nueva hoja de consulta"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdEliminar"
            Object.ToolTipText     =   "Elimina una hoja de consulta"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdLimpiar"
            Object.ToolTipText     =   "Limpia toda la consulta"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdAscendente"
            Object.ToolTipText     =   "Ordenar columna ascendente"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDescendente"
            Object.ToolTipText     =   "Ordenar columna descendente"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdAbrir"
            Object.ToolTipText     =   "Abre un archivo sql"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdGrabar"
            Object.ToolTipText     =   "Graba el contenido de hoja de consulta"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdImprimir"
            Object.ToolTipText     =   "Imprime el sql activo"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdCopiar"
            Object.ToolTipText     =   "Copia el contenido al portapapeles"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdCortar"
            Object.ToolTipText     =   "Corta el contenido al portapapeles"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdPegar"
            Object.ToolTipText     =   "Pega el contenido del portapapeles "
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdUndo"
            Object.ToolTipText     =   "Deshacer cambios realizados"
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdBuscar"
            Object.ToolTipText     =   "Ejecutar sentencia SQL ..."
            Object.Tag             =   ""
            ImageIndex      =   30
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDetener"
            Object.ToolTipText     =   "Detener ejecución sentencia SQL ..."
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Sale de Easy Query!"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   8655
      TabIndex        =   2
      Top             =   465
      Visible         =   0   'False
      Width           =   8715
      Begin VB.PictureBox picResource 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   3
         Left            =   420
         Picture         =   "QueryIt.frx":08CA
         ScaleHeight     =   150
         ScaleWidth      =   1050
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   780
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.PictureBox picResource 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   2
         Left            =   0
         Picture         =   "QueryIt.frx":0D03
         ScaleHeight     =   720
         ScaleWidth      =   9180
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   9180
      End
   End
   Begin SysTray.XSystray cSystray1 
      Left            =   4095
      Top             =   2805
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "QueryIt.frx":25A3
      TrayTip         =   "SysTray Control."
   End
   Begin ComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12383
            Text            =   "Easy Query ! Chile 2001"
            TextSave        =   "Easy Query ! Chile 2001"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "10:12 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   3495
      Top             =   615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   30
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":28BD
            Key             =   ""
            Object.Tag             =   "B&uscar"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":2A33
            Key             =   ""
            Object.Tag             =   "&Copiar"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":2BA9
            Key             =   ""
            Object.Tag             =   "C&ortar"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":2D1F
            Key             =   ""
            Object.Tag             =   "&Pegar"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":2E95
            Key             =   ""
            Object.Tag             =   "&Nuevo"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":300B
            Key             =   ""
            Object.Tag             =   "&Guardar"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":3181
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":32F7
            Key             =   ""
            Object.Tag             =   "&Limpiar"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":346D
            Key             =   ""
            Object.Tag             =   "&Abrir"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":357F
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":36F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":3A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":3D99
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":40EB
            Key             =   "cmdBorrar"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":41FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":454F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":48A1
            Key             =   ""
            Object.Tag             =   "&Skins"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":4BBB
            Key             =   ""
            Object.Tag             =   "&Eliminar"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":4ED5
            Key             =   ""
            Object.Tag             =   "&Conectar"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":51EF
            Key             =   ""
            Object.Tag             =   "&Desconectar"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":5509
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":561B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":572D
            Key             =   ""
            Object.Tag             =   "&Deshacer"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":583F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":5A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":5BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":5DCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":60E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":62C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "QueryIt.frx":649B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivox 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuQuery_Conectar 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu mnuQuery_Desconectar 
         Caption         =   "&Desconectar"
      End
      Begin VB.Menu mnuQuery_Sep99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Nuevahoja 
         Caption         =   "&Nuevo formulario de consultas"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuQuery_EliminarHoja 
         Caption         =   "&Eliminar formulario de consultas"
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
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdicion_Pegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdicion_Cortar 
         Caption         =   "C&ortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdicion_Borrar 
         Caption         =   "&Borrar"
         Shortcut        =   ^B
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
      Begin VB.Menu mnuQuery_CopiarDatos 
         Caption         =   "&Copiar registros seleccionados"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuQuery_sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_SeleccionarTodo 
         Caption         =   "&Seleccionar todos y copiar"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuQuery_DesSeleccionarTodo 
         Caption         =   "&DesSeleccionar todos los registros"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuQuery_sep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Imprimir 
         Caption         =   "&Imprimir resultado de la búsqueda"
      End
      Begin VB.Menu mnuQuery_CopiarHeader 
         Caption         =   "Copiar &header de consulta"
      End
      Begin VB.Menu mnuQuery_EiminarFila 
         Caption         =   "&Eliminar fila de consulta"
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
      Begin VB.Menu sex2 
         Caption         =   "-"
      End
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
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuArchivo_Restaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu mnuArchivo_Acercade 
         Caption         =   "&Acerca de Easy Query ! ..."
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_Configuracion 
         Caption         =   "C&onfiguración"
      End
      Begin VB.Menu mnuOpciones_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Colores 
         Caption         =   "&Colores"
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

Public m_Detener As Boolean

Private TrxActiva As Boolean

'Private WithEvents HelpObj As HelpCallBack

Dim IsInFocus As Boolean
Dim xad As Long

Dim Largo As Integer

Private WithEvents xConnection As ADODB.Connection
Attribute xConnection.VB_VarHelpID = -1

'Private rsPaso As New ADODB.Recordset
Private Sub AbreSql()

    On Local Error GoTo ErrorAbreSql
    
    Dim cc As New GCommonDialog
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim Sql As String
    Dim Glosa As String
    
    Glosa = "Archivos de comandos Sql (*.SQL)|*.SQL|"
    Glosa = Glosa & "Archivos de comandos Texto (*.TXT)|*.TXT|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , App.Path, "Abrir archivo ...", "SQL", Me.hwnd)) Then
       Exit Sub
    End If
            
    If Archivo <> "" Then
        nArchivo = FreeFile
        
        Open Archivo For Input As #nArchivo
            Sql = Input(LOF(nArchivo), nArchivo)
        Close #nArchivo
        frmMain.ActiveForm.txtQuery.Text = vbNullString
        frmMain.ActiveForm.txtQuery.SelColor = QBColor(1)
        frmMain.ActiveForm.txtQuery.Text = Sql
    End If
    
    'frmMain.ActiveForm.txtQuery.Visible = False
    frmMain.ActiveForm.HayCambios = True
                
    Call frmMain.ActiveForm.FormateaSentencias
    
    'frmMain.ActiveForm.txtQuery.Visible = True
    frmMain.ActiveForm.txtQuery.SetFocus
    
    Set cc = Nothing
    
    GoTo SalirAbreSql
    
ErrorAbreSql:
    Resume SalirAbreSql
    
SalirAbreSql:
    Err = 0
    
End Sub



Private Sub Ascendente()

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
            
End Sub

Private Function ChequeaComandoCritico() As Boolean

    Dim k As Integer
    Dim ret As Boolean
    
    Dim Sql As String
    
    Sql = Trim$(frmMain.ActiveForm.txtQuery.Text)
    
    ret = False
    
    For k = 1 To UBound(Comandos)
        If Comandos(k).Activo Then
            If UCase$(Comandos(k).Comando) = UCase$(Left$(Sql, Len(Comandos(k).Comando))) Then
                ret = True
                Exit For
            End If
        End If
    Next k
    
    ChequeaComandoCritico = ret
    
End Function

Private Sub Descendente()

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
    Dim msg As String
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then Exit Sub

    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
            
    ActiveConexion = ConexionActiva(ci.Text)
    
    If cState(ActiveConexion).Trx = True Then
        
        msg = "Confirma finalizar transacción activa."
        If Confirma(msg) = vbNo Then Exit Sub
        
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

    Dim k As Integer
    Dim qR As Integer
    Dim Sql As String
    Dim ArrayCount As Integer
    Dim i As Integer
    Dim f As Form
    Dim NComandos
    'Dim sComando
    Dim Activo As String
    
    qR = 1
    
    Call Hourglass(hwnd, True)
    
    Call GrabaIni(C_INI, "skins", "skin", glbPathSkin)
    
    If UBound(Document) = 0 Then Exit Sub
    
    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = ArrayCount To 1 Step -1
        If Not fState(i).Deleted Then
            Set f = Document(i)
            Sql = f.txtQuery.Text
            If Sql <> "" Then
                Call f.txtQuery.SaveFile(App.Path & "\query" & qR & ".sql", 1)
                qR = qR + 1
            End If
            Unload f
        End If
    Next
    
    Call GrabaIni(C_INI, "Querys", "nquerys", CStr(qR - 1))
    
    ReDim Document(0)
    
    Call GrabaIni(C_INI, "ArchivosMDB", "nArchivosMdb", UBound(aArchivosMdb))
        
    For k = 1 To UBound(aArchivosMdb)
        Call GrabaIni(C_INI, "ArchivosMDB", "Archivo" & k, aArchivosMdb(k))
    Next k
    
    'grabar comandos
    
    'leer comandos criticos
    NComandos = UBound(Comandos)
         
    If NComandos > 0 Then
        For k = 1 To NComandos
            If Comandos(k).Activo Then
                Activo = "1"
            Else
                Activo = "0"
            End If
            
            Call GrabaIni(C_INI, "Comandos", "comando" & k, Comandos(k).Comando & "," & Activo)
        Next k
    End If
    
    Call GrabaIni(C_INI, "Comandos", "ncomandos", NComandos)
    
    Call Hourglass(hwnd, False)
    
End Sub




Private Sub GrabaSQL()

    On Local Error GoTo ErrorGrabaSQL
    
    Dim cc As New GCommonDialog
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim Sql As String
    Dim Glosa As String
        
    Glosa = "Archivos de comandos Sql (*.SQL)|*.SQL|"
    Glosa = Glosa & "Archivos de comandos Texto (*.TXT)|*.TXT|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar como ...", "SQL", Me.hwnd)) Then
       Exit Sub
    End If
           
    If Archivo <> "" Then
        nArchivo = FreeFile
        Sql = frmMain.ActiveForm.txtQuery.Text
        Archivo = Left$(Archivo, Len(Archivo) - 1)
        Open Archivo & ".sql" For Output As #nArchivo
            Print #nArchivo, Sql
        Close #nArchivo
    End If
    
    Set cc = Nothing
    
    GoTo SalirGrabaSQL
    
ErrorGrabaSQL:
    Resume SalirGrabaSQL
    
SalirGrabaSQL:
    Err = 0
    
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

    Dim k As Integer
    Dim Sql As String
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
    
    glbTimeOut = LeeIni("General", "timeout", C_INI)
    If glbTimeOut = "" Then glbTimeOut = "0"
    
    Nquerys = LeeIni("Querys", "nquerys", C_INI)
    If Nquerys = "" Then Nquerys = 0
     
    If Nquerys > 0 Then
        For k = 1 To Nquerys
        
            On Error Resume Next
            
            nFreeFile = FreeFile
            Sql = vbNullString
            Open App.Path & "\query" & k & ".sql" For Input As #nFreeFile
                Sql = Input$(LOF(nFreeFile), nFreeFile)
            Close #nFreeFile
                        
            If Sql <> "" Then
                If k > 1 Then
                    Call FileNew
                Else
                    fState(1).Tag = "q1"
                    fState(1).Dirty = 1
                End If
                Document(k).txtQuery.Text = Sql
                Document(k).Tag = "q" & k
            End If
            
            Err = 0
            
        Next k
    End If
    
    Nsentencias = LeeIni("Sentencias", "nSentencias", C_INI)
    If Nsentencias = "" Then Nsentencias = 0
    
    ReDim aSentencias(Val(Nsentencias))
    
    For k = 1 To Nsentencias
        Sentencia = LeeIni("Sentencias", "Sentencia" & k, C_INI)
        Color = LeeIni("Sentencias", "Color" & k, C_INI)
        aSentencias(k).Glosa = Sentencia
        aSentencias(k).Color = CLng(Color)
    Next k
      
    nArchivosMdb = LeeIni("ArchivosMDB", "nArchivosMdb", C_INI)
    If nArchivosMdb = "" Then nArchivosMdb = 0
    
    ReDim aArchivosMdb(Val(nArchivosMdb))
    
    For k = 1 To nArchivosMdb
        aArchivosMdb(k) = LeeIni("ArchivosMDB", "Archivo" & k, C_INI)
    Next k
    
    'leer comandos criticos
    NComandos = LeeIni("Comandos", "ncomandos", C_INI)
    If NComandos = "" Then NComandos = 0
     
    If NComandos > 0 Then
        c = 1
        For k = 1 To NComandos
            sComando = LeeIni("Comandos", "comando" & k, C_INI)
            
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
        Next k
    End If
    
End Sub


Private Sub Limpiar()

    On Local Error Resume Next
    frmMain.ActiveForm.griQuery.MaxRows = 1
    frmMain.ActiveForm.griQuery.MaxCols = 1
    
    Call frmMain.ActiveForm.griQuery.SetText(1, 0, "")
    Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
    Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
    frmMain.ActiveForm.txtQuery.Text = ""
    frmMain.ActiveForm.txtQuery.SetFocus
    Err = 0
            
End Sub

Private Sub PegarTexto()

    On Local Error Resume Next
            
    frmMain.ActiveForm.txtQuery.SelStart = 1
    frmMain.ActiveForm.txtQuery.SelLength = Len(frmMain.ActiveForm.txtQuery.Text)
    frmMain.ActiveForm.txtQuery.SelColor = QBColor(1)
    frmMain.ActiveForm.txtQuery.SelStart = 1
    frmMain.ActiveForm.txtQuery.SelLength = 0
    frmMain.ActiveForm.txtQuery.SelColor = QBColor(1)
    Call Pegar(frmMain.ActiveForm.txtQuery.hwnd)
    frmMain.ActiveForm.HayCambios = True
    'frmMain.ActiveForm.txtQuery.Visible = False
    Call frmMain.ActiveForm.FormateaSentencias
    'frmMain.ActiveForm.txtQuery.Visible = True
    Err = 0
            
End Sub

Private Sub RollTRX()
    
    On Local Error Resume Next
    
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim msg As String
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then Exit Sub

    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem

    ActiveConexion = ConexionActiva(ci.Text)
    
    If cState(ActiveConexion).Trx = True Then
        
        msg = "Confirma cancelar transacción activa."
        If Confirma(msg) = vbNo Then Exit Sub
        
        cState(ActiveConexion).Trx = False
        DBConnection(ActiveConexion).RollbackTrans
        
        Toolbar.Buttons("cmdIniTrx").Enabled = True
        Toolbar.Buttons("cmdFinTrx").Enabled = False
        Toolbar.Buttons("cmdRollTrx").Enabled = False
    End If
    Err = 0
            
    Set ci = Nothing
    
End Sub

Private Function TipoDeCampo(ByVal Tipo As DataTypeEnum) As String

    Dim ret As String
        
    If Tipo = adBigInt Then ret = "BigInt"
    If Tipo = adBinary Then ret = "Binary"
    If Tipo = adBoolean Then ret = "Boolean"
    If Tipo = adBSTR Then ret = "BSTR"
    If Tipo = adChar Then ret = "Char"
    If Tipo = adCurrency Then ret = "Currency"
    If Tipo = adDate Then ret = "Date"
    If Tipo = adDBDate Then ret = "Date"
    If Tipo = adDBTime Then ret = "Time"
    If Tipo = adDBTimeStamp Then ret = "TimeStamp"
    If Tipo = adDecimal Then ret = "Decimal"
    If Tipo = adDouble Then ret = "Double"
    If Tipo = adEmpty Then ret = "Empty"
    If Tipo = adError Then ret = "Error"
    If Tipo = adGUID Then ret = "GUID"
    If Tipo = adIDispatch Then ret = "IDispatch"
    If Tipo = adInteger Then ret = "Integer"
    If Tipo = adIUnknown Then ret = "Unknown"
    If Tipo = adLongVarBinary Then ret = "LongVarBinary"
    If Tipo = adLongVarChar Then ret = "LongVarChar"
    If Tipo = adLongVarWChar Then ret = "LongVarWChar"
    If Tipo = adNumeric Then ret = "Numeric"
    If Tipo = adSingle Then ret = "Single"
    If Tipo = adSmallInt Then ret = "SmallInt"
    If Tipo = adTinyInt Then ret = "TinyInt"
    If Tipo = adUnsignedBigInt Then ret = "UnsignedBigInt"
    If Tipo = adUnsignedInt Then ret = "UnsignedInt"
    If Tipo = adUnsignedSmallInt Then ret = "UnsignedSmallInt"
    If Tipo = adUnsignedTinyInt Then ret = "UnsignedTinyInt"
    If Tipo = adUserDefined Then ret = "UserDefined"
    If Tipo = adVarBinary Then ret = "VarBinary"
    If Tipo = adVarChar Then ret = "VarChar"
    If Tipo = adVariant Then ret = "Variant"
    If Tipo = adVarWChar Then ret = "VarWChar"
    If Tipo = adWChar Then ret = "WChar"
        
    TipoDeCampo = ret
    
End Function

Private Sub Undo()

    On Local Error Resume Next
    Call Deshacer(frmMain.ActiveForm.txtQuery.hwnd)
    frmMain.ActiveForm.HayCambios = True
    'frmMain.ActiveForm.txtQuery.Visible = False
    Call frmMain.ActiveForm.FormateaSentencias
    'frmMain.ActiveForm.txtQuery.Visible = True
    Err = 0
            
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)

    If Button = vbLeftButton Then
        frmMain.Visible = True
        frmMain.WindowState = vbNormal
        frmMain.Show
    End If
    
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)

    If Button = vbRightButton Then
        PopupMenu mnuArchivo
    End If
    
End Sub



Private Sub HelpObj_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)

    If Enabled Then
        staBar.Panels(1).Text = MenuHelp$
    Else
        staBar.Panels(1).Text = ""
    End If
  
End Sub

Private Sub MDIForm_Load()
    
    Dim k As Integer
    Dim op
    'Dim intCount As Integer
       
    Cargando = True
    
    glbPathSkin = LeeIni("skins", "skin", C_INI)
        
    On Local Error Resume Next
    
    glbPathSkin = ""
    
    Set m_cN = New cNeoCaption
          
    'SkinForm.SkinPath = glbPathSkin
        
    Err = 0
    
    Call Hourglass(hwnd, True)
    
    'Show
   
    gbLastPos = 0
    
    ReDim Document(1)
    ReDim fState(1)
    
    ReDim DBConnection(0)
    ReDim cState(0)
    ReDim aArchivosMdb(0)
    ReDim Comandos(0)
    
    Document(1).Tag = "q1"
    
    fState(1).Dirty = False
    fState(k).Tag = "q1"
    fState(k).Dirty = 1
    
    Conectado = False
    TrxActiva = False
    
    Call LeeQuerys
    
    'Call InstalaSistemaAyuda
    
    Set cSystray1.TrayIcon = Me.Icon
    cSystray1.TrayTip = App.Title & " - Creado por Luis Núñez Ibarra."
            
    App.HelpFile = App.Path & "\symphonyx.hlp"
    
    Largo = 0
        
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
    
    'Call HookForm(Me)
    
    Call Hourglass(hwnd, False)
    
End Sub

Public Sub EjecutaQuery()

    On Local Error GoTo ErrorEjecutaQuery
    
    Dim itmx As ListItem
    
    Dim msg As String
    Dim Sql As String
    'Dim SqlPaso As String
    Dim Campo As Field
    Dim Ncampos As Integer
    Dim Nombre
    Dim k As Long
    Dim f As Integer
    Dim e As Integer
    Dim r As Long
    Dim cx As Integer
    Dim sCampo As String
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    Dim Flagx As Boolean
    Dim VarPaso As String
    Dim nDecimal As Integer
    
    If Not Conectado Then
        MsgBox "Debe conectarse primero.", vbCritical
        Exit Sub
    End If
    
    If frmMain.ActiveForm Is Nothing Then
        MsgBox "Debe activar un formulario o agregar una ventana de consultas.", vbCritical
        Exit Sub
    End If
    
    Sql = Trim$(frmMain.ActiveForm.txtQuery.Text)
    
    If Sql = "" Then
        MsgBox "Debe digitar una sentencia sql.", vbCritical
        Exit Sub
    End If
    
    If frmMain.ActiveForm.m_Ejecutando Then Exit Sub
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then
        MsgBox "Seleccione una conexión.", vbCritical
        frmMain.ActiveForm.imgConexiones.SetFocus
        Exit Sub
    End If
    
    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    
    ActiveConexion = ConexionActiva(ci.Text)
        
    'frmMain.ActiveForm.pgbQuery.Total = 0
    
    frmMain.ActiveForm.griQuery.MaxRows = 1
    frmMain.ActiveForm.griQuery.MaxCols = 1
    
    Call frmMain.ActiveForm.griQuery.SetText(1, 0, "Selec")
    Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
    Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
    Call frmMain.ActiveForm.lviewDetalle.ListItems.Clear
    
    Call Hourglass(hwnd, True)
    
    Toolbar.Buttons("cmdDetener").Enabled = True
            
    If ChequeaComandoCritico() Then
        msg = "Confirma ejecutar este comando crítico definido."
        If Confirma(msg) = vbNo Then Exit Sub
    End If
    
    frmMain.ActiveForm.m_Ejecutando = True
    
    If frmMain.ActiveForm.Rs.State <> 0 Then
        frmMain.ActiveForm.Rs.Close
    End If
    
    If Left$(UCase$(Sql), 6) = "SELECT" Then
            
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando Select ...."
        
        Set xConnection = DBConnection(ActiveConexion)
        xConnection.CommandTimeout = glbTimeOut
        
        Set frmMain.ActiveForm.Rs.ActiveConnection = xConnection
        Set frmMain.ActiveForm.Rs = New ADODB.Recordset
        
        frmMain.ActiveForm.Rs.CursorLocation = adUseServer
        
        frmMain.ActiveForm.Rs.Open Sql, xConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        Ncampos = 2
        
        If Not frmMain.ActiveForm.Rs.EOF Then
            Toolbar.Buttons("cmdSalir").Enabled = False
            'frmMain.ActiveForm.Rs.MoveLast
            
            'auxTotalRegistros = frmMain.ActiveForm.Rs.RecordCount
            
            'frmMain.ActiveForm.Rs.MoveFirst
            
            'frmMain.ActiveForm.pgbQuery.Visible = True
            'frmMain.ActiveForm.pgbQuery.Total = auxTotalRegistros
            
            'frmMain.ActiveForm.pgbQuery.ZOrder 0
            'frmMain.ActiveForm.Refresh
            
            'frmMain.ActiveForm.griQuery.MaxRows = auxTotalRegistros
            
            frmMain.ActiveForm.TabQuery.Tabs(2).Selected = True
            
            With frmMain.ActiveForm.Rs
                'cargar headers de la spread
                For Each Campo In .Fields
                    If Ncampos > frmMain.ActiveForm.griQuery.MaxCols Then
                        frmMain.ActiveForm.griQuery.MaxCols = frmMain.ActiveForm.griQuery.MaxCols + 1
                    End If
                    Nombre = Campo.Name
                    Set itmx = frmMain.ActiveForm.lviewCampos.ListItems.Add()
                    itmx.Text = Nombre
        
                    frmMain.ActiveForm.griQuery.Col = Ncampos
                    Call frmMain.ActiveForm.griQuery.SetText(Ncampos, 0, Nombre)
                    
                    sCampo = TipoDeCampo(Campo.Type)
                    Flagx = False
                    
                    Select Case sCampo
                        Case "BigInt", "Currency", "Decimal", "Double", "Integer", "LongVarBinary", _
                             "Numeric", "Single", "SmallInt", "TinyInt", "UnsignedBigInt", _
                            "UnsignedInt", "UnsignedSmallInt", "UnsignedTinyInt"
                        
                            Flagx = True
                            frmMain.ActiveForm.griQuery.Row = -1
                            frmMain.ActiveForm.griQuery.Row2 = -1
                            frmMain.ActiveForm.griQuery.Col = Ncampos
                            frmMain.ActiveForm.griQuery.Col2 = Ncampos
                            frmMain.ActiveForm.griQuery.CellType = CellTypeFloat
                            frmMain.ActiveForm.griQuery.TypeHAlign = 1
                            frmMain.ActiveForm.griQuery.TypeVAlign = 2
                            
                            If Campo.NumericScale = 0 Then
                                frmMain.ActiveForm.griQuery.TypeFloatMin = "-" & String$(Campo.Precision, "9")
                                frmMain.ActiveForm.griQuery.TypeFloatMax = String$(Campo.Precision, "9")
                            Else
                                If Campo.NumericScale <> 255 Then
                                    frmMain.ActiveForm.griQuery.TypeFloatMin = "-" & String$(Campo.Precision - Campo.NumericScale, "9") & "," & String$(Campo.NumericScale, "9")
                                    frmMain.ActiveForm.griQuery.TypeFloatMax = String$(Campo.Precision - Campo.NumericScale, "9") & "," & String$(Campo.NumericScale, "9")
                                Else
                                    frmMain.ActiveForm.griQuery.TypeFloatMin = "-" & String$(Campo.Precision, "9")
                                    frmMain.ActiveForm.griQuery.TypeFloatMax = String$(Campo.Precision, "9")
                                End If
                            End If
                            
                            frmMain.ActiveForm.griQuery.TypeFloatMoney = False
                            frmMain.ActiveForm.griQuery.TypeFloatSeparator = True
                            frmMain.ActiveForm.griQuery.TypeFloatDecimalPlaces = Campo.NumericScale
                            frmMain.ActiveForm.griQuery.TypeFloatCurrencyChar = Asc("$")
                            frmMain.ActiveForm.griQuery.TypeFloatSepChar = Asc(".")
                            frmMain.ActiveForm.griQuery.TypeFloatDecimalChar = Asc(",")
                    End Select
                    
                    Ncampos = Ncampos + 1
                Next
            
                'llenar listview de campos
                For Each Campo In .Fields
                    Nombre = Campo.Name
                    Set itmx = frmMain.ActiveForm.lviewDetalle.ListItems.Add()
                    itmx.Text = Nombre
                    itmx.SubItems(1) = TipoDeCampo(Campo.Type)
                    
                    If Left$(itmx.SubItems(1), 1) = "C" Then
                        itmx.SubItems(2) = Campo.ActualSize
                        itmx.SubItems(3) = "0"
                    Else
                        If Campo.NumericScale = 0 Then
                            itmx.SubItems(2) = Campo.Precision
                            itmx.SubItems(3) = 0
                        Else
                            If Campo.NumericScale <> 255 Then
                                itmx.SubItems(2) = Campo.Precision - Campo.NumericScale
                                itmx.SubItems(3) = Campo.NumericScale
                            Else
                                itmx.SubItems(2) = Campo.Precision
                                itmx.SubItems(3) = 0
                            End If
                        End If
                    End If
                                        
                    If Campo.Attributes = adFldIsNullable Then
                        itmx.SubItems(4) = "Si"
                    Else
                        itmx.SubItems(4) = "No"
                    End If
                Next
                
                'cargar datos en la spread
                f = 1
                frmMain.ActiveForm.griQuery.Redraw = False
                Do While Not .EOF
                    e = DoEvents()
                    If frmMain.ActiveForm.m_Detener Then
                        frmMain.ActiveForm.m_Detener = False
                        Exit Do
                    End If
                    If frmMain.ActiveForm.griQuery.DataRowCnt >= frmMain.ActiveForm.griQuery.MaxRows Then
                        frmMain.ActiveForm.griQuery.MaxRows = frmMain.ActiveForm.griQuery.MaxRows + 1
                    End If
                    
                    'llenar x columnas
                    For k = 2 To Ncampos - 1
                        If Not IsNull(.Fields(k - 2)) Then
                            If IsNumeric(.Fields(k - 2)) Then
                                VarPaso = .Fields(k - 2)
                                If InStr(VarPaso, ",") Then
                                    VarPaso = Mid$(VarPaso, InStr(1, VarPaso, ",") + 1)
                                    
                                    If CLng(VarPaso) = 0 Then
                                        frmMain.ActiveForm.griQuery.Col = k
                                        frmMain.ActiveForm.griQuery.Row = f
                                        frmMain.ActiveForm.griQuery.Col2 = k
                                        frmMain.ActiveForm.griQuery.Row2 = f
                                        frmMain.ActiveForm.griQuery.TypeFloatDecimalPlaces = 0
                                    Else
                                        frmMain.ActiveForm.griQuery.Col = k
                                        frmMain.ActiveForm.griQuery.Row = f
                                        frmMain.ActiveForm.griQuery.Col2 = k
                                        frmMain.ActiveForm.griQuery.Row2 = f
                                        frmMain.ActiveForm.griQuery.TypeFloatDecimalPlaces = Len(VarPaso)
                                    End If
                                Else
                                    frmMain.ActiveForm.griQuery.Col = k
                                    frmMain.ActiveForm.griQuery.Row = f
                                    frmMain.ActiveForm.griQuery.Col2 = k
                                    frmMain.ActiveForm.griQuery.Row2 = f
                                    frmMain.ActiveForm.griQuery.TypeFloatDecimalPlaces = 0
                                End If
                            End If
                            
                            Call frmMain.ActiveForm.griQuery.SetText(k, f, .Fields(k - 2))
                        End If
                    Next k
                    f = f + 1
                    .MoveNext
                Loop
                frmMain.ActiveForm.griQuery.Redraw = True
            End With
        End If
        
        frmMain.ActiveForm.griQuery.Row = 1
        frmMain.ActiveForm.Rs.Close
        
        Set frmMain.ActiveForm.Rs = Nothing
        
        frmMain.ActiveForm.staQuery.Panels(2).Text = CStr(frmMain.ActiveForm.griQuery.DataRowCnt) & " registros."
    Else
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando instrucción SQL ...."
        
        Set xConnection = DBConnection(ActiveConexion)
        xConnection.CommandTimeout = glbTimeOut
        
        Call xConnection.Execute(Sql)
        MsgBox "Acción ejecutada con éxito!", vbInformation
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Listo."
    End If
                    
    Call Hourglass(hwnd, False)
    
    frmMain.ActiveForm.m_Ejecutando = False
    
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.Row = -1
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row2 = frmMain.ActiveForm.griQuery.DataRowCnt
    frmMain.ActiveForm.griQuery.BlockMode = True
    frmMain.ActiveForm.griQuery.Protect = False
        
    'frmMain.ActiveForm.pgbQuery.Visible = False
    'FrmMain.ActiveForm.pgbQuery.Valor = 0
    
    GoTo Salir
    
ErrorEjecutaQuery:
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.staQuery.Panels(2).Text = Error$
    End If
    
    Resume Salir
    
Salir:
    Err = 0
    Call Hourglass(hwnd, False)
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.m_Ejecutando = False
    End If
    
    If frmMain.ActiveForm.Rs.State <> 0 Then
        frmMain.ActiveForm.Rs.Close
    End If
        
    Toolbar.Buttons("cmdSalir").Enabled = True
    
    Set Campo = Nothing
    Set itmx = Nothing
    Set ci = Nothing
        
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim msg As String
    
    If Not Toolbar.Buttons("cmdSalir").Enabled Then
        Cancel = 1
        Exit Sub
    End If
    
    msg = "Confirma salir de " & App.Title
    
    If Confirma(msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    cSystray1.InTray = False
    DoEvents
    
    Call GrabaQuerys
    
    Unload frmTablas
    Unload frmColores
    Unload frmBrowser
    
    DoEvents
    
    Me.Hide
    
    frmExit.Show vbModal
    
End Sub
Private Sub MDIForm_Resize()

    If Me.WindowState = vbMinimized Then
        cSystray1.InTray = True
        
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = False
        If mnuOpciones_FTexto.Checked = True Then frmColores.Visible = False
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = False
        'Timer1.Enabled = True
        
        Me.Visible = False
    Else
        cSystray1.InTray = False
        Set cSystray1.TrayIcon = Me.Icon
        
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = True
        If mnuOpciones_FTexto.Checked = True Then frmColores.Visible = True
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = True
        
        Me.Visible = True
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
        
    If glbPathSkin = "" Then
        'Call Uninstall(Me.hwnd)
    End If
    
    'Call UnHookForm(Me)
    
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
                
    End
    
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
    ShellExecute Me.hwnd, vbNullString, "http://proyectovb.topcities.com/", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0

End Sub

Private Sub mnuEdicion_Borrar_Click()
    frmMain.ActiveForm.txtQuery.Text = ""
End Sub

Private Sub mnuEdicion_Buscar_Click()
    frmFind.Show
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
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdUndo"))
End Sub

Private Sub mnuEdicion_Pegar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdPegar"))
End Sub

Private Sub mnuEdicion_Reemplazar_Click()
    frmReemplazar.Show vbModal
End Sub

Private Sub mnuEdicion_SeleccionarTodo_Click()
    
    On Local Error Resume Next
    
    frmMain.ActiveForm.txtQuery.SelStart = 0
    frmMain.ActiveForm.txtQuery.SelLength = Len(frmMain.ActiveForm.txtQuery.Text)
    frmMain.ActiveForm.txtQuery.SetFocus
    
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

Private Sub mnuOpciones_FTexto_Click()

    If Not Cargando Then
        mnuOpciones_FTexto.Checked = Not mnuOpciones_FTexto.Checked
        
        If mnuOpciones_FTexto.Checked Then
            frmColores.Show
        Else
            Unload frmColores
        End If
    End If
    
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
    
    Dim k As Integer
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
    
    For k = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.GetText(1, k, Sel)
        If Sel = "1" Then
            frmMain.ActiveForm.griQuery.Col = 2
            frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
            frmMain.ActiveForm.griQuery.Row = k
            frmMain.ActiveForm.griQuery.Row2 = k
            gsBuffer = gsBuffer & frmMain.ActiveForm.griQuery.Clip
        End If
    Next k
    
    Clipboard.Clear
    Clipboard.SetText gsBuffer, 1
    
    Call Hourglass(hwnd, False)
    
    Err = 0
    
End Sub

Private Sub mnuQuery_CopiarHeader_Click()

    Dim k As Integer
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    Clipboard.Clear
    
    frmMain.ActiveForm.griQuery.Row = 0
    
    For k = 2 To frmMain.ActiveForm.griQuery.MaxCols
        frmMain.ActiveForm.griQuery.Col = k
        gsBuffer = gsBuffer & frmMain.ActiveForm.griQuery.Text & " , "
    Next k
    
    gsBuffer = Left$(gsBuffer, Len(gsBuffer) - 3)
    
    Clipboard.SetText gsBuffer, 1
    
End Sub

Private Sub mnuQuery_Desconectar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdDesconectar"))
End Sub

Private Sub mnuQuery_DesSeleccionarTodo_Click()

    On Local Error Resume Next
    
    Dim k As Integer
    'Dim Sel
    
    Call Hourglass(hwnd, True)
    
    For k = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.SetText(1, k, "")
    Next k
    
    Call Hourglass(hwnd, False)

    Err = 0
    
End Sub

Private Sub mnuQuery_EiminarFila_Click()
        
    Dim k As Long
    Dim Sel
    
    If frmMain.ActiveForm.griQuery.ActiveRow > 0 Then
        For k = frmMain.ActiveForm.griQuery.MaxRows To 1 Step -1
            Call frmMain.ActiveForm.griQuery.GetText(1, k, Sel)
            If Sel = "1" Then
                frmMain.ActiveForm.griQuery.Row = k
                frmMain.ActiveForm.griQuery.Row2 = k
                frmMain.ActiveForm.griQuery.BlockMode = True
                frmMain.ActiveForm.griQuery.Action = ActionDeleteRow
                frmMain.ActiveForm.griQuery.BlockMode = False
            End If
        Next k
    End If
    
End Sub

Private Sub mnuQuery_EliminarHoja_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdEliminar"))
End Sub

Private Sub mnuQuery_ExportarCSV_Click()

    Dim Archivo As String
    
    Archivo = SaveDialog(hwnd, "Archivos de Sql (*.CSV)|*.CSV", "Guardar archivo como ...", App.Path)
    
    If StripNulls(Archivo) <> "" Then
        If InStr(Archivo, ".") = 0 Then
            Archivo = Archivo & ".csv"
        End If
        'exportar a csv
        If CSVExport(Archivo) Then
            MsgBox "Archivo exportado con éxito!", vbInformation
        End If
    End If
    
End Sub

Private Sub mnuQuery_ExportarHTM_Click()

    Dim Archivo As String
    
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
    
End Sub

Private Sub mnuQuery_ExportarTXT_Click()

    Dim Archivo As String
    
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
    
End Sub

Private Sub mnuQuery_ExportarXLS_Click()

    Dim Archivo As String
    
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
    
End Sub

Private Sub mnuQuery_ExportarXML_Click()

    Dim Archivo As String
    
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
    
End Sub


Private Sub mnuQuery_GuardarSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdGrabar"))
End Sub

Private Sub mnuQuery_Imprimir_Click()
    
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.ColHidden = True
    
    frmPreview.lHwnd = frmMain.ActiveForm.griQuery.hwnd
    frmPreview.Show vbModal
        
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.ColHidden = False
    
End Sub

Private Sub mnuQuery_ImprimirSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdImprimir"))
End Sub

Private Sub mnuQuery_Nuevahoja_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdNuevo"))
End Sub


Private Sub mnuQuery_Salir_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdSalir"))
End Sub


Private Sub mnuQuery_SeleccionarTodo_Click()

    On Local Error Resume Next
    
    Dim k As Integer
    'Dim Sel
    
    Call Hourglass(hwnd, True)
    
    For k = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.SetText(1, k, "1")
    Next k
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    Clipboard.Clear
    
    frmMain.ActiveForm.griQuery.Col = 2
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row = -1
    frmMain.ActiveForm.griQuery.Row2 = -1
            
    gsBuffer = frmMain.ActiveForm.griQuery.Clip
    
    Clipboard.SetText gsBuffer, 1
    
    Call Hourglass(hwnd, False)
    
    Err = 0
    
End Sub

Private Sub mnuQuery_VerTabla_Click()
    frmVerTabla.Show vbModal
End Sub

Private Sub mnuQuery_VerTodasTablas_Click()
    mnuQuery_VerTodasTablas.Checked = Not mnuQuery_VerTodasTablas.Checked
    
    If mnuQuery_VerTodasTablas.Checked = True Then
        frmTablas.Show
    Else
        Unload frmTablas
    End If
    
End Sub

Private Sub mnuTablas_ImprimirCampos_Click()

    On Local Error GoTo ErrorImprimirCampos
    
    Dim itmx As ListItem
    Dim Header As String
    
    Dim k As Integer
    Dim j As Integer
    
    Dim Campos As Integer
    
    Dim Buffer As String
    
    Buffer = ""
    
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
        For k = 1 To .ListItems.Count
            Set itmx = .ListItems(k)
            
            Buffer = itmx.Text & vbTab
            
            For j = 1 To 4
                If j = 3 Then
                    Buffer = Buffer & Space$(5 - Len(itmx.SubItems(j))) & itmx.SubItems(j) & vbTab
                ElseIf j = 4 Then
                    Buffer = Buffer & Space$(9 - Len(itmx.SubItems(j))) & itmx.SubItems(j) & vbTab
                ElseIf j = 5 Then
                    Buffer = Buffer & Space$(5 - Len(itmx.SubItems(j))) & itmx.SubItems(j)
                Else
                    Buffer = Buffer & itmx.SubItems(j) & vbTab
                End If
            Next j
            
            Printer.Print Buffer
        Next k
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
    frmMain.Arrange vbCascade
End Sub

Private Sub mnuVentana_Horizontal_Click()
    frmMain.Arrange vbTileHorizontal
End Sub


Private Sub mnuVentana_Organizar_Click()
    frmMain.Arrange vbArrangeIcons
End Sub




Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)

    'Dim k As Integer
    Dim msg As String
    Dim fIndex
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    
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
            Call FileNew(fIndex)
            
            If mnuOpciones_Browser.Checked = True Then
                frmBrowser.treQuerys.Nodes.Add("EJ1", tvwChild, _
                frmMain.ActiveForm.Tag, frmMain.ActiveForm.Caption, 1).EnsureVisible
            End If
        Case "cmdEliminar"
            Call EliminarHojaConsulta
            
        Case "cmdLimpiar"
            Call Limpiar
        Case "cmdAscendente"
            Call Ascendente
        Case "cmdDescendente"
            Call Descendente
        Case "cmdAbrir"
            Call AbreSql
        Case "cmdGrabar"
            Call GrabaSQL
        Case "cmdImprimir"
            Call Imprimir
        Case "cmdCopiar"
            On Local Error Resume Next
            Call Copiar(frmMain.ActiveForm!txtQuery.hwnd)
            Err = 0
        Case "cmdCortar"
            On Local Error Resume Next
            Call Cortar(frmMain.ActiveForm!txtQuery.hwnd)
            Err = 0
        Case "cmdPegar"
            Call PegarTexto
        Case "cmdUndo"
            Call Undo
        Case "cmdBuscar"
            Call EjecutaQuery
        Case "cmdDetener"
            frmMain.ActiveForm.m_Detener = True
        Case "cmdSalir"
            'Call GrabaQuerys
            Unload Me
    End Select
    
End Sub

Private Sub xConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
        
    Dim msg As String
    
    If pError Is Nothing Then Exit Sub
    
    If pError.Number <> 0 Then
        
        msg = pError.Source & vbNewLine & vbNewLine
        msg = msg & pError.Description & vbNewLine
        MsgBox msg, vbCritical
        
    End If
    
End Sub

Private Sub xConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    
    Dim e As Integer
    
    Do While adStatus <> adStatusOK
        e = DoEvents()
    Loop
    
End Sub



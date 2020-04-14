VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3559BCB7-0351-11D5-B855-444553540001}#1.0#0"; "XSystray.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SymphonyX"
   ClientHeight    =   3930
   ClientLeft      =   75
   ClientTop       =   2205
   ClientWidth     =   12060
   HelpContextID   =   20
   Icon            =   "SymphonyX.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   767
      ButtonWidth     =   688
      ButtonHeight    =   661
      AllowCustomize  =   0   'False
      ImageList       =   "imgList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   27
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdConectar"
            Object.ToolTipText     =   "Conecta hacia fuente de datos"
            Object.Tag             =   ""
            ImageIndex      =   35
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDesconectar"
            Object.ToolTipText     =   "Desconecta una conexión previa"
            Object.Tag             =   ""
            ImageIndex      =   36
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
            ImageIndex      =   17
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdFinTrx"
            Object.ToolTipText     =   "Finalizar Transacción"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdRollTrx"
            Object.ToolTipText     =   "Cancelar Transaccion"
            Object.Tag             =   ""
            ImageIndex      =   16
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
            ImageIndex      =   25
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdLimpiar"
            Object.ToolTipText     =   "Limpia toda la consulta"
            Object.Tag             =   ""
            ImageIndex      =   11
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
            ImageIndex      =   45
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDescendente"
            Object.ToolTipText     =   "Ordenar columna descendente"
            Object.Tag             =   ""
            ImageIndex      =   46
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
            ImageIndex      =   12
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
            ImageIndex      =   13
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
            Object.ToolTipText     =   "Pega el contenido del portapapeles hacia hoja activa"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdUndo"
            Object.ToolTipText     =   "Deshacer cambios realizados"
            Object.Tag             =   ""
            ImageIndex      =   47
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
            Object.ToolTipText     =   "Busca según contenido de hoja activa"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cmdDetener"
            Object.ToolTipText     =   "Detiene consulta de hoja activa"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Sale de SymphonyX!"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin SysTray.XSystray cSystray1 
      Left            =   1440
      Top             =   3000
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "SymphonyX.frx":08CA
      TrayTip         =   "SysTray Control."
   End
   Begin MSComctlLib.ImageList imgIconos 
      Left            =   1470
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SymphonyX.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SymphonyX.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SymphonyX.frx":17DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "SymphonyX.frx":1AF8
      TabIndex        =   2
      Top             =   435
      Width           =   480
   End
   Begin ComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3645
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18203
            Text            =   "SymphonyX - Creado por Luis Núñez Ibarra "
            TextSave        =   "SymphonyX - Creado por Luis Núñez Ibarra "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "15:16"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   180
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   47
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":1BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":1D45
            Key             =   ""
            Object.Tag             =   "&Copiar"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":1EBB
            Key             =   ""
            Object.Tag             =   "&Cortar"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2031
            Key             =   ""
            Object.Tag             =   "&Pegar"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":21A7
            Key             =   ""
            Object.Tag             =   "&Nuevo"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":231D
            Key             =   ""
            Object.Tag             =   "&Guardar"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2493
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2609
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":277F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":28F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2A6B
            Key             =   ""
            Object.Tag             =   "&Limpiar"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2BE1
            Key             =   ""
            Object.Tag             =   "&Abrir"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2CF3
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":2E69
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":31BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":350D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":385F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":3BB1
            Key             =   "cmdBorrar"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":3CC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":4015
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":4367
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":46B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":47CB
            Key             =   ""
            Object.Tag             =   "&Skins"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":4AE5
            Key             =   ""
            Object.Tag             =   "&Colores"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":4DFF
            Key             =   ""
            Object.Tag             =   "&Eliminar"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":5119
            Key             =   ""
            Object.Tag             =   "&Organizar"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":5433
            Key             =   ""
            Object.Tag             =   "&Cascada"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":574D
            Key             =   ""
            Object.Tag             =   "&Horizontal"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":5A67
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":5C41
            Key             =   ""
            Object.Tag             =   "&Configurar"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":5F5B
            Key             =   ""
            Object.Tag             =   "&Seleccionar"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":6275
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":644F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":6629
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":6803
            Key             =   ""
            Object.Tag             =   "&Conectar"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":6B1D
            Key             =   ""
            Object.Tag             =   "&Desconectar"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":6E37
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":7011
            Key             =   ""
            Object.Tag             =   "&DesSeleccionar"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":732B
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":7645
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":795F
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":7C79
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":7F93
            Key             =   ""
            Object.Tag             =   "&Registros"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":82AD
            Key             =   ""
            Object.Tag             =   "&Todo"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":8487
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":8599
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SymphonyX.frx":86AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivox 
      Caption         =   "&Archivo"
      HelpContextID   =   100
      Begin VB.Menu mnuQuery_Conectar 
         Caption         =   "&Conectar "
         HelpContextID   =   20
      End
      Begin VB.Menu mnuQuery_Desconectar 
         Caption         =   "&Desconectar"
         HelpContextID   =   30
      End
      Begin VB.Menu mnuQuery_Sep99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Nuevahoja 
         Caption         =   "&Nuevo formulario de consultas"
         HelpContextID   =   40
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuQuery_EliminarHoja 
         Caption         =   "&Eliminar formulario de consultas"
         HelpContextID   =   50
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuQuery_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_AbrirconsultaSQL 
         Caption         =   "&Abrir consulta SQL"
         HelpContextID   =   120
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuQuery_GuardarSQL 
         Caption         =   "&Guardar SQL"
         HelpContextID   =   140
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuQuery_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_ConfigImpresora 
         Caption         =   "&Configurar Página"
         HelpContextID   =   150
      End
      Begin VB.Menu mnuQuery_ImprimirSQL 
         Caption         =   "&Imprimir SQL"
         HelpContextID   =   160
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuQuery_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Salir 
         Caption         =   "&Salir"
         HelpContextID   =   170
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      HelpContextID   =   300
      Begin VB.Menu mnuEdicion_Deshacer 
         Caption         =   "&Deshacer"
         HelpContextID   =   310
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEdicion_Copiar 
         Caption         =   "&Copiar"
         HelpContextID   =   320
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdicion_Pegar 
         Caption         =   "&Pegar"
         HelpContextID   =   330
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdicion_Cortar 
         Caption         =   "&Cortar"
         HelpContextID   =   340
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdicion_Borrar 
         Caption         =   "&Borrar"
         HelpContextID   =   350
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdicion_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_SeleccionarTodo 
         Caption         =   "&Seleccionar todo el texto"
         HelpContextID   =   350
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
      Enabled         =   0   'False
      HelpContextID   =   370
      Begin VB.Menu mnuQuery_CopiarDatos 
         Caption         =   "&Copiar registros seleccionados"
         HelpContextID   =   380
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuQuery_sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_SeleccionarTodo 
         Caption         =   "&Seleccionar todos y copiar"
         HelpContextID   =   390
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuQuery_DesSeleccionarTodo 
         Caption         =   "&DesSeleccionar todos los registros"
         HelpContextID   =   400
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuQuery_sep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_Imprimir 
         Caption         =   "Imprimir resultado de la búsqueda"
         HelpContextID   =   410
      End
      Begin VB.Menu mnuQuery_CopiarHeader 
         Caption         =   "Copiar header de consulta"
         HelpContextID   =   420
      End
      Begin VB.Menu sex1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuery_VerTabla 
         Caption         =   "Ver estructura de tabla"
         HelpContextID   =   430
      End
      Begin VB.Menu mnuQuery_VerTodasTablas 
         Caption         =   "Ver todas las tablas de conexión"
         HelpContextID   =   440
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      HelpContextID   =   340
      Visible         =   0   'False
      Begin VB.Menu mnuArchivo_Restaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu mnuArchivo_Acercade 
         Caption         =   "&Acerca de SymphonyX ..."
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
      HelpContextID   =   450
      Begin VB.Menu mnuOpciones_Colores 
         Caption         =   "&Colores"
         HelpContextID   =   470
      End
      Begin VB.Menu mnuOpciones_Skin 
         Caption         =   "&Skin"
         HelpContextID   =   480
      End
      Begin VB.Menu mnuOpciones_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Browser 
         Caption         =   "&Browser de Querys"
         HelpContextID   =   490
      End
      Begin VB.Menu mnuOpciones_FTexto 
         Caption         =   "&Formato de Texto"
         HelpContextID   =   500
      End
      Begin VB.Menu mnuOpciones_SiempreVisible 
         Caption         =   "Siempre &visible"
         HelpContextID   =   510
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      HelpContextID   =   360
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
      HelpContextID   =   370
      Begin VB.Menu mnuAyuda_Contenido 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "&Indice"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAyuda_Busqueda 
         Caption         =   "&Búsqueda"
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_Web 
         Caption         =   "&SymphonyX en el WEB"
      End
      Begin VB.Menu mnuAyuda_Email 
         Caption         =   "&Email a ProyectoVB"
      End
      Begin VB.Menu mnuAyuda_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_AcercaDe 
         Caption         =   "&Acerca de SymphonyX..."
      End
   End
   Begin VB.Menu mnuTablas 
      Caption         =   "Tablas"
      HelpContextID   =   380
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

'Private NumQuery As Integer
Private QueryActiva As Integer
Private TrxActiva As Boolean

'Private WithEvents HelpObj As HelpCallBack

Dim IsInFocus As Boolean
Dim xad As Long

Dim Largo As Integer
Dim nTitulo As Integer
Dim Titulo1 As String
Dim Titulo2 As String
Dim Titulo3 As String
Dim Titulo4 As String
'Implements ISubclass
'Private m_emr As EMsgResponse

Dim nIcono As Integer

Private Sub AbreSql()

    On Local Error GoTo ErrorAbreSql
    
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim Sql As String
    
    Archivo = OpenDialog(hWnd, "*.sql", "Abrir SQL ...", App.Path)
    
    If Archivo <> "" Then
        nArchivo = FreeFile
        
        Open Archivo For Input As #nArchivo
            Sql = Input(LOF(nArchivo), nArchivo)
        Close #nArchivo
        frmMain.ActiveForm.txtQuery.Text = Sql
    End If
    
    Call frmMain.ActiveForm.FormateaSentencias
    
    GoTo SalirAbreSql
    
ErrorAbreSql:
    Resume SalirAbreSql
    
SalirAbreSql:
    Err = 0
    
End Sub



Public Function ConexionActiva(ByVal Conexion As String) As Integer

    Dim K As Integer
    
    Dim ret As Integer
    
    For K = 1 To UBound(cState)
        If cState(K).Deleted = False Then
            If cState(K).Conexion = Conexion Then
                ret = K
                Exit For
            ElseIf cState(K).LlaveMdb = Conexion Then
                ret = K
                Exit For
            End If
        End If
    Next K
    
    ConexionActiva = ret
    
End Function

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

Public Sub GrabaQuerys()

    Dim K As Integer
    Dim qR As Integer
    Dim Sql As String
    Dim ArrayCount As Integer
    Dim I As Integer
    Dim F As Form
    
    qR = 1
    
    Call Hourglass(hWnd, True)
    
    Call GrabaIni(C_INI, "skins", "skin", glbPathSkin)
    
    If UBound(Document) = 0 Then Exit Sub
    
    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For I = 1 To ArrayCount
        If Not fState(I).Deleted Then
            Set F = Document(I)
            Sql = F.txtQuery.Text
            If Sql <> "" Then
                Call GrabaIni(C_INI, "Querys", "query" & qR, ClearEnterInString(Sql))
                qR = qR + 1
            End If
            Unload F
        End If
    Next
    
    Call GrabaIni(C_INI, "Querys", "nquerys", CStr(qR - 1))
    
    ReDim Document(0)
    
    Call GrabaIni(C_INI, "ArchivosMDB", "nArchivosMdb", UBound(aArchivosMdb))
        
    For K = 1 To UBound(aArchivosMdb)
        Call GrabaIni(C_INI, "ArchivosMDB", "Archivo" & K, aArchivosMdb(K))
    Next K
    
    Call Hourglass(hWnd, False)
    
End Sub




Private Sub GrabaSQL()

    On Local Error GoTo ErrorGrabaSQL
    
    Dim Archivo As String
    Dim nArchivo As Integer
    Dim Sql As String
    
    Archivo = SaveDialog(hWnd, "*.sql", "Guardar como ...", App.Path)
    
    If Archivo <> "" Then
        nArchivo = FreeFile
        Sql = frmMain.ActiveForm.txtQuery.Text
        Open Archivo For Output As #nArchivo
            Print #nArchivo, Sql
        Close #nArchivo
    End If
    
    GoTo SalirGrabaSQL
    
ErrorGrabaSQL:
    Resume SalirGrabaSQL
    
SalirGrabaSQL:
    Err = 0
    
End Sub


Private Sub Imprimir()

    On Local Error GoTo ErrorImprimir
    
    If frmMain.ActiveForm.txtQuery.Text = "" Then Exit Sub
        
    Call Hourglass(hWnd, True)
    
    If Not ShowPrinter(Me) Then GoTo SalirImprimir
    
    Printer.Print frmMain.ActiveForm.txtQuery.Text
        
    GoTo SalirImprimir
    
ErrorImprimir:
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hWnd, False)
    Err = 0
    
End Sub

Private Sub LeeQuerys()

    Dim K As Integer
    Dim Sql As String
    Dim Sentencia As String
    Dim Color
    Dim Nquerys
    Dim Nsentencias
    Dim nArchivosMdb
    
    glbTimeOut = LeeIni("General", "timeout", C_INI)
    If glbTimeOut = "" Then glbTimeOut = "0"
    
    Nquerys = LeeIni("Querys", "nquerys", C_INI)
    If Nquerys = "" Then Nquerys = 0
     
    If Nquerys > 0 Then
        For K = 1 To Nquerys
            Sql = LeeIni("Querys", "query" & K, C_INI)
            
            If Sql <> "" Then
                If K > 1 Then
                    Call FileNew
                Else
                    fState(1).Tag = "q1"
                    fState(1).Dirty = 1
                End If
                Document(K).txtQuery.Text = Sql
                Document(K).Tag = "q" & K
            End If
        Next K
    End If
    
    Nsentencias = LeeIni("Sentencias", "nSentencias", C_INI)
    If Nsentencias = "" Then Nsentencias = 0
    
    ReDim aSentencias(Val(Nsentencias))
    
    For K = 1 To Nsentencias
        Sentencia = LeeIni("Sentencias", "Sentencia" & K, C_INI)
        Color = LeeIni("Sentencias", "Color" & K, C_INI)
        aSentencias(K).Glosa = Sentencia
        aSentencias(K).Color = CLng(Color)
    Next K
      
    nArchivosMdb = LeeIni("ArchivosMDB", "nArchivosMdb", C_INI)
    If nArchivosMdb = "" Then nArchivosMdb = 0
    
    ReDim aArchivosMdb(Val(nArchivosMdb))
    
    For K = 1 To nArchivosMdb
        aArchivosMdb(K) = LeeIni("ArchivosMDB", "Archivo" & K, C_INI)
    Next K
    
End Sub


Private Function TipoDeCampo(ByVal Tipo As Integer) As String

    Dim ret As String
    
    Select Case Tipo
        'Case adArray: ret = "Array"
        Case adBigInt: ret = "BigInt"
        Case adBinary: ret = "Binary"
        Case adBoolean: ret = "Boolean"
        'Case adByRef: ret = "ByRef"
        Case adBSTR: ret = "BSTR"
        Case adChar: ret = "Char"
        Case adCurrency: ret = "Currency"
        Case adDate: ret = "Date"
        Case adDBDate: ret = "Date"
        Case adDBTime: ret = "Time"
        Case adDBTimeStamp: ret = "TimeStamp"
        Case adDecimal: ret = "Decimal"
        Case adDouble: ret = "Double"
        Case adEmpty: ret = "Empty"
        Case adError: ret = "Error"
        Case adGUID: ret = "GUID"
        Case adIDispatch: ret = "IDispatch"
        Case adInteger: ret = "Integer"
        Case adIUnknown: ret = "Unknown"
        Case adLongVarBinary: ret = "LongVarBinary"
        Case adLongVarChar: ret = "LongVarChar"
        Case adLongVarWChar: ret = "LongVarWChar"
        Case adNumeric: ret = "Numeric"
        Case adSingle: ret = "Single"
        Case adSmallInt: ret = "SmallInt"
        Case adTinyInt: ret = "TinyInt"
        Case adUnsignedBigInt: ret = "UnsignedBigInt"
        Case adUnsignedInt: ret = "UnsignedInt"
        Case adUnsignedSmallInt: ret = "UnsignedSmallInt"
        Case adUnsignedTinyInt: ret = "UnsignedTinyInt"
        Case adUserDefined: ret = "UserDefined"
        Case adVarBinary: ret = "VarBinary"
        Case adVarChar: ret = "VarChar"
        Case adVariant: ret = "Variant"
        'Case adVector: ret = "Vector"
        Case adVarWChar: ret = "VarWChar"
        Case adWChar: ret = "WChar"
    End Select
    
    TipoDeCampo = ret
    
End Function

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)

    If Button = vbLeftButton Then
        frmMain.Visible = True
        frmMain.WindowState = vbMaximized
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
        staBar.Panels(1).Text = App.Title
    End If
  
End Sub

Private Sub MDIForm_Load()
    
    Dim K As Integer
    Dim op
    Dim intCount As Integer
       
    Cargando = True
    
    glbPathSkin = LeeIni("skins", "skin", C_INI)
    
    'If glbPathSkin = "" Then glbPathSkin = App.Path
        
    On Local Error Resume Next
    
    SkinForm.SkinPath = glbPathSkin
    
    Err = 0
    
    Call Hourglass(hWnd, True)
    
    'Show
   
    ReDim Document(1)
    ReDim fState(1)
    
    ReDim DBConnection(0)
    ReDim cState(0)
    ReDim aArchivosMdb(0)
    
    Document(1).Tag = "q1"
    
    fState(1).Dirty = False
    fState(K).Tag = "q1"
    fState(K).Dirty = 1
    
    Conectado = False
    TrxActiva = False
    
    Call LeeQuerys
    
    'Set HelpObj = New HelpCallBack

    'Call Install(Me.hwnd, HelpObj, imgList, True, True)
  
    'FontName Me.hwnd, "Tahoma"
    'FontSize Me.hwnd, 8&
  
    Set cSystray1.TrayIcon = Me.Icon
    cSystray1.TrayTip = App.Title & " - Creado por Luis Núñez Ibarra."
            
    Largo = 0
    nTitulo = 1
        
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
    
    SetAppHelp hWnd
    
    Call Hourglass(hWnd, False)
    
End Sub

Public Sub EjecutaQuery()

    On Local Error GoTo ErrorEjecutaQuery
    
    Dim itmx As ListItem
    
    Dim Msg As String
    Dim Sql As String
    Dim SqlPaso As String
    Dim Campo As Field
    Dim Ncampos As Integer
    Dim Nombre
    Dim K As Long
    Dim F As Integer
    Dim e As Integer
    Dim r As Long
    Dim auxTotalRegistros
    Dim sCampo As String
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    
    If Not Conectado Then
        MsgBox "Debe conectarse primero.", vbCritical
        Exit Sub
    End If
    
    Sql = Trim$(frmMain.ActiveForm.txtQuery.Text)
    
    If Sql = "" Then Exit Sub
    
    If frmMain.ActiveForm.m_Ejecutando Then Exit Sub
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then
        MsgBox "Seleccione una conexión.", vbCritical
        Exit Sub
    End If
    
    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    
    ActiveConexion = ConexionActiva(ci.Text)
        
    frmMain.ActiveForm.m_Ejecutando = True
    
    frmMain.ActiveForm.pgbQuery.Total = 0
    
    frmMain.ActiveForm.griQuery.MaxRows = 1
    frmMain.ActiveForm.griQuery.MaxCols = 1
    
    Call frmMain.ActiveForm.griQuery.SetText(1, 0, "Selec")
    Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
    Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
    Call frmMain.ActiveForm.lviewDetalle.ListItems.Clear
    
    Call Hourglass(hWnd, True)
    
    Toolbar.Buttons("cmdDetener").Enabled = True
            
    If Left$(UCase$(Sql), 6) = "SELECT" Then
            
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando Select ...."
        
        Set frmMain.ActiveForm.Rs = New ADODB.Recordset
        frmMain.ActiveForm.Rs.CursorLocation = adUseClient
        frmMain.ActiveForm.Rs.Open Sql, DBConnection(ActiveConexion), adOpenStatic
        
        Ncampos = 2
        
        If Not frmMain.ActiveForm.Rs.EOF Then
            
            frmMain.ActiveForm.Rs.MoveLast
            
            auxTotalRegistros = frmMain.ActiveForm.Rs.RecordCount
            
            frmMain.ActiveForm.Rs.MoveFirst
            
            frmMain.ActiveForm.pgbQuery.Visible = True
            frmMain.ActiveForm.pgbQuery.Total = auxTotalRegistros
            
            frmMain.ActiveForm.pgbQuery.ZOrder 0
            frmMain.ActiveForm.Refresh
            
            frmMain.ActiveForm.griQuery.MaxRows = auxTotalRegistros
            
            frmMain.ActiveForm.TabQuery.Tabs(2).Selected = True
            
            With frmMain.ActiveForm.Rs
                For Each Campo In .Fields
                    If Ncampos > frmMain.ActiveForm.griQuery.MaxCols Then
                        frmMain.ActiveForm.griQuery.MaxCols = frmMain.ActiveForm.griQuery.MaxCols + 1
                    End If
                    Nombre = Campo.Name
                    Set itmx = frmMain.ActiveForm.lviewCampos.ListItems.Add()
                    itmx.Text = Nombre
        
                    Call frmMain.ActiveForm.griQuery.SetText(Ncampos, 0, Nombre)
                    
                    sCampo = TipoDeCampo(Campo.Type)
                    
                    Select Case sCampo
                        Case "BigInt", "Currency", "Decimal", "Double", "Integer", "LongVarBinary", _
                             "Numeric", "Single", "SmallInt", "TinyInt", "UnsignedBigInt", _
                            "UnsignedInt", "UnsignedSmallInt", "UnsignedTinyInt"
                        
                            frmMain.ActiveForm.griQuery.Row = -1
                            frmMain.ActiveForm.griQuery.Col = Ncampos
                            frmMain.ActiveForm.griQuery.CellType = CellTypeFloat
                            frmMain.ActiveForm.griQuery.TypeHAlign = 1
                            frmMain.ActiveForm.griQuery.TypeFloatMin = "-999999999999"
                            frmMain.ActiveForm.griQuery.TypeFloatMax = "9999999999999"
                            frmMain.ActiveForm.griQuery.TypeFloatMoney = False
                            frmMain.ActiveForm.griQuery.TypeFloatSeparator = True
                            frmMain.ActiveForm.griQuery.TypeFloatDecimalPlaces = 0
                            frmMain.ActiveForm.griQuery.TypeFloatCurrencyChar = Asc("$")
                            frmMain.ActiveForm.griQuery.TypeFloatSepChar = Asc(".")
                            frmMain.ActiveForm.griQuery.TypeFloatDecimalChar = Asc(",")
                    End Select
                    
                    Ncampos = Ncampos + 1
                Next
            
                For Each Campo In .Fields
                    Nombre = Campo.Name
                    Set itmx = frmMain.ActiveForm.lviewDetalle.ListItems.Add()
                    itmx.Text = Nombre
                    itmx.SubItems(1) = TipoDeCampo(Campo.Type)
                    itmx.SubItems(2) = Campo.DefinedSize
                    If Campo.Attributes = adFldIsNullable Then
                        itmx.SubItems(3) = "Si"
                    Else
                        itmx.SubItems(3) = "No"
                    End If
                Next
                
                F = 1
                Do While Not .EOF
                    e = DoEvents()
                    If frmMain.ActiveForm.m_Detener Then
                        frmMain.ActiveForm.m_Detener = False
                        Exit Do
                    End If
                    If frmMain.ActiveForm.griQuery.DataRowCnt >= frmMain.ActiveForm.griQuery.MaxRows Then
                        frmMain.ActiveForm.griQuery.MaxRows = frmMain.ActiveForm.griQuery.MaxRows + 1
                    End If
                    For K = 2 To Ncampos - 1
                        Call frmMain.ActiveForm.griQuery.SetText(K, F, .Fields(K - 2))
                    Next K
                    frmMain.ActiveForm.pgbQuery.Valor = F
                    frmMain.ActiveForm.pgbQuery.Progress
                    F = F + 1
                    .MoveNext
                Loop
            End With
        End If
        
        frmMain.ActiveForm.griQuery.Row = 1
        frmMain.ActiveForm.Rs.Close
        
        frmMain.ActiveForm.staQuery.Panels(2).Text = CStr(frmMain.ActiveForm.griQuery.DataRowCnt) & " registros."
    Else
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Ejecutando instrucción SQL ...."
        DBConnection(ActiveConexion).CommandTimeout = glbTimeOut
        Call DBConnection(ActiveConexion).Execute(Sql)
        MsgBox "Acción ejecutada con éxito!", vbInformation
        frmMain.ActiveForm.staQuery.Panels(2).Text = "Listo."
    End If
                    
    Call Hourglass(hWnd, False)
    
    frmMain.ActiveForm.m_Ejecutando = False
    
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.Row = -1
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row2 = frmMain.ActiveForm.griQuery.DataRowCnt
    frmMain.ActiveForm.griQuery.BlockMode = True
    frmMain.ActiveForm.griQuery.Protect = False
        
    frmMain.ActiveForm.pgbQuery.Visible = False
    frmMain.ActiveForm.pgbQuery.Valor = 0
    
    GoTo Salir
    
ErrorEjecutaQuery:
    frmMain.ActiveForm.staQuery.Panels(2).Text = Error$
    MsgBox "EjecutaQuery : " & Err & " " & Error$, vbCritical
    Err = 0
    Resume Salir
    
Salir:
    'Timer1.Enabled = True
    Call Hourglass(hWnd, False)
    frmMain.ActiveForm.m_Ejecutando = False
    Set Campo = Nothing
    Set itmx = Nothing
    Set ci = Nothing
    
    'Unload frmBuscar
    
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim Msg As String
    
    Msg = "Confirma salir de " & App.Title
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    cSystray1.InTray = False
    DoEvents
    
    Call GrabaQuerys
    
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState = vbMinimized Then
        cSystray1.InTray = True
        
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = False
        If mnuOpciones_FTexto.Checked = True Then frmColores.Visible = False
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = False
        nIcono = 1
        'Timer1.Enabled = True
        
        Me.Visible = False
    Else
        cSystray1.InTray = False
        Set cSystray1.TrayIcon = Me.Icon
        
        If mnuOpciones_Browser.Checked = True Then frmBrowser.Visible = True
        If mnuOpciones_FTexto.Checked = True Then frmColores.Visible = True
        If mnuQuery_VerTodasTablas.Checked = True Then frmTablas.Visible = True
        
        'Timer1.Enabled = False
        Me.Visible = True
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
        
    'Call Uninstall(Me.hwnd)
  
    'Set HelpObj = Nothing
    
    QuitHelp
    
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
    
    Unload frmTablas
    Unload frmColores
    Unload frmBrowser
    
    DoEvents
    
End Sub

Private Sub mnuArchivo_Acercade_Click()
    frmAcercaDe.Show vbModal
End Sub

Private Sub mnuArchivo_Restaurar_Click()

    frmMain.Visible = True
    frmMain.WindowState = vbMaximized
    frmMain.Show
    
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuAyuda_AcercaDe_Click()
    frmAcercaDe.Show vbModal
End Sub

Private Sub mnuAyuda_Busqueda_Click()
     Call SearchHelp
End Sub

Private Sub mnuAyuda_Contenido_Click()
    Call ShowHelpContents
End Sub

Private Sub mnuAyuda_Email_Click()
    Shell_Email
End Sub

Private Sub mnuAyuda_Indice_Click()
    Call ShowHelpOnHelp
End Sub

Private Sub mnuAyuda_Web_Click()
    Shell_PaginaWeb
End Sub

Private Sub mnuEdicion_Borrar_Click()
    frmMain.ActiveForm.txtQuery.Text = ""
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
            Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
        Else
            Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
        End If
    End If
    
End Sub

Private Sub mnuOpciones_Skin_Click()
   frmConfSkin.Show vbModal
End Sub


Private Sub mnuQuery_AbrirconsultaSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdAbrir"))
End Sub

Private Sub mnuQuery_Conectar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdConectar"))
End Sub

Private Sub mnuQuery_ConfigImpresora_Click()
    
    Call ShowPageSetupDlg(hWnd)
        
End Sub

Private Sub mnuQuery_CopiarDatos_Click()

    On Local Error Resume Next
    
    Dim K As Integer
    Dim Sel
    
    Call Hourglass(hWnd, True)
    
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
    
    Call Hourglass(hWnd, False)
    
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
    
    gsBuffer = Left$(gsBuffer, Len(gsBuffer) - 3)
    
    Clipboard.SetText gsBuffer, 1
    
End Sub

Private Sub mnuQuery_Desconectar_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdDesconectar"))
End Sub

Private Sub mnuQuery_DesSeleccionarTodo_Click()

    On Local Error Resume Next
    
    Dim K As Integer
    Dim Sel
    
    Call Hourglass(hWnd, True)
    
    For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.SetText(1, K, "")
    Next K
    
    Call Hourglass(hWnd, False)

    Err = 0
    
End Sub

Private Sub mnuQuery_EliminarHoja_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdEliminar"))
End Sub

Private Sub mnuQuery_GuardarSQL_Click()
    Call Toolbar_ButtonClick(Toolbar.Buttons("cmdGrabar"))
End Sub

Private Sub mnuQuery_Imprimir_Click()
    
    frmMain.ActiveForm.griQuery.Col = 1
    frmMain.ActiveForm.griQuery.ColHidden = True
    
    frmPreview.lHwnd = frmMain.ActiveForm.griQuery.hWnd
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
    
    Dim K As Integer
    Dim Sel
    
    Call Hourglass(hWnd, True)
    
    For K = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.SetText(1, K, "1")
    Next K
    
    gsBuffer = ""
        
    frmMain.ActiveForm.griQuery.Clip = gsBuffer
    
    Clipboard.Clear
    
    frmMain.ActiveForm.griQuery.Col = 2
    frmMain.ActiveForm.griQuery.Col2 = frmMain.ActiveForm.griQuery.MaxCols
    frmMain.ActiveForm.griQuery.Row = -1
    frmMain.ActiveForm.griQuery.Row2 = -1
            
    gsBuffer = frmMain.ActiveForm.griQuery.Clip
    
    Clipboard.SetText gsBuffer, 1
    
    Call Hourglass(hWnd, False)
    
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
    
    Dim K As Integer
    Dim j As Integer
    
    Dim Campos As Integer
    
    Dim Buffer As String
    
    Buffer = ""
    
    Campos = frmMain.ActiveForm.griQuery.MaxCols
        
    Printer.Font = "Courier New"
    Printer.Font.Size = 10
    
    Printer.Print "Nombre Físico" & vbTab & "Tipo" & vbTab & "Largo" & vbTab & "Acepta Nulos"
    
    With frmMain.ActiveForm.lviewDetalle
        For K = 1 To .ListItems.Count
            Set itmx = .ListItems(K)
            
            Buffer = itmx.Text & vbTab
            
            For j = 1 To 3
                Buffer = Buffer & itmx.SubItems(j) & vbTab
            Next j
            
            Printer.Print Buffer
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
    frmMain.Arrange vbCascade
End Sub

Private Sub mnuVentana_Horizontal_Click()
    frmMain.Arrange vbTileHorizontal
End Sub


Private Sub mnuVentana_Organizar_Click()
    frmMain.Arrange vbArrangeIcons
End Sub




Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)

    Dim K As Integer
    Dim Msg As String
    Dim fIndex
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    
    Select Case Button.Key
        Case "cmdConectar"
            frmConexiones.Show vbModal
        Case "cmdDesconectar"
            frmConexiones.Show vbModal
        Case "cmdIniTrx"
            
            On Local Error Resume Next
            
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
            Err = 0
        Case "cmdFinTrx"
            
            On Local Error Resume Next
            
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
            Err = 0
        Case "cmdRollTrx"
            
            On Local Error Resume Next
            
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
            
        Case "cmdNuevo"
            Call FileNew(fIndex)
            
            If mnuOpciones_Browser.Checked = True Then
                frmBrowser.treQuerys.Nodes.Add("EJ1", tvwChild, _
                frmMain.ActiveForm.Tag, frmMain.ActiveForm.Caption, 1).EnsureVisible
            End If
        Case "cmdEliminar"
            Call EliminarHojaConsulta
            
        Case "cmdLimpiar"
            On Local Error Resume Next
            frmMain.ActiveForm.griQuery.MaxRows = 1
            frmMain.ActiveForm.griQuery.MaxCols = 1
            
            Call frmMain.ActiveForm.griQuery.SetText(1, 0, "")
            Call frmMain.ActiveForm.griQuery.SetText(1, 1, "")
            Call frmMain.ActiveForm.lviewCampos.ListItems.Clear
            frmMain.ActiveForm.txtQuery.Text = ""
            frmMain.ActiveForm.txtQuery.SetFocus
            Err = 0
            
        Case "cmdAscendente"
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
        Case "cmdDescendente"
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
        Case "cmdAbrir"
            Call AbreSql
        Case "cmdGrabar"
            Call GrabaSQL
        Case "cmdImprimir"
            Call Imprimir
        Case "cmdCopiar"
            On Local Error Resume Next
            Call Copiar(frmMain.ActiveForm!txtQuery.hWnd)
            Err = 0
        Case "cmdCortar"
            On Local Error Resume Next
            Call Cortar(frmMain.ActiveForm!txtQuery.hWnd)
            Err = 0
        Case "cmdPegar"
            On Local Error Resume Next
            
            frmMain.ActiveForm.txtQuery.SelStart = 1
            frmMain.ActiveForm.txtQuery.SelLength = Len(frmMain.ActiveForm.txtQuery.Text)
            frmMain.ActiveForm.txtQuery.SelColor = QBColor(1)
            frmMain.ActiveForm.txtQuery.SelStart = 1
            frmMain.ActiveForm.txtQuery.SelLength = 0
            frmMain.ActiveForm.txtQuery.SelColor = QBColor(1)
            Call Pegar(frmMain.ActiveForm.txtQuery.hWnd)
            frmMain.ActiveForm.HayCambios = True
            Call frmMain.ActiveForm.FormateaSentencias
            Err = 0
        Case "cmdUndo"
            On Local Error Resume Next
            Call Deshacer(frmMain.ActiveForm.txtQuery.hWnd)
            frmMain.ActiveForm.HayCambios = True
            Call frmMain.ActiveForm.FormateaSentencias
            Err = 0
        Case "cmdBuscar"
            Call EjecutaQuery
        Case "cmdDetener"
            frmMain.ActiveForm.m_Detener = True
        Case "cmdSalir"
            'Call GrabaQuerys
            Unload Me
    End Select
    
End Sub


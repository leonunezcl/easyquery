VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConexiones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conectar ..."
   ClientHeight    =   6780
   ClientLeft      =   2940
   ClientTop       =   2235
   ClientWidth     =   6870
   Icon            =   "Conexiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6705
      Left            =   0
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   21
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Remover MDB"
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
      Index           =   7
      Left            =   5295
      TabIndex        =   20
      ToolTipText     =   "Eliminar base de datos access de la lista"
      Top             =   3975
      Width           =   1515
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo base de datos MDB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   435
      TabIndex        =   17
      Top             =   4695
      Width           =   4815
      Begin VB.OptionButton opt2000 
         Caption         =   "Access 2000"
         Height          =   255
         Left            =   2700
         TabIndex        =   19
         ToolTipText     =   "Abrir base de datos Access 2000"
         Top             =   360
         Width           =   1485
      End
      Begin VB.OptionButton opt97 
         Caption         =   "Access 97"
         Height          =   255
         Left            =   510
         TabIndex        =   18
         ToolTipText     =   "Abrir base de datos access 97"
         Top             =   360
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Conectar MDB"
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
      Index           =   6
      Left            =   5295
      TabIndex        =   16
      ToolTipText     =   "Conectar a base de datos access"
      Top             =   3015
      Width           =   1515
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Desconecta MDB"
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
      Index           =   5
      Left            =   5295
      TabIndex        =   15
      ToolTipText     =   "Cerrar base de datos access"
      Top             =   3495
      Width           =   1515
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Agregar MDB"
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
      Index           =   2
      Left            =   5295
      TabIndex        =   14
      ToolTipText     =   "Agregar base de datos access a lista"
      Top             =   2535
      Width           =   1515
   End
   Begin VB.Frame fraMDB 
      Caption         =   "Bases de datos Access (.MDB) (97/2000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   405
      TabIndex        =   12
      Top             =   2415
      Width           =   4815
      Begin MSComctlLib.TreeView treeMdb 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Doble clic para abrir base de datos access"
         Top             =   270
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3201
         _Version        =   393217
         Indentation     =   229
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         SingleSel       =   -1  'True
         ImageList       =   "imgLst"
         Appearance      =   1
      End
   End
   Begin VB.Frame fraODBC 
      Caption         =   "Origen de Datos ODBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   405
      TabIndex        =   10
      Top             =   45
      Width           =   4815
      Begin MSComctlLib.TreeView treConexiones 
         Height          =   1905
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Doble clic para conectar a origen de datos ODBC"
         Top             =   300
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3360
         _Version        =   393217
         Indentation     =   229
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         SingleSel       =   -1  'True
         ImageList       =   "imgLst"
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Agregar ODBC"
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
      Index           =   1
      Left            =   5295
      TabIndex        =   9
      ToolTipText     =   "Agregar fuente de datos ODBC"
      Top             =   465
      Width           =   1515
   End
   Begin VB.Frame fraUsuario 
      Caption         =   "Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   405
      TabIndex        =   5
      Top             =   5595
      Width           =   4845
      Begin VB.TextBox txtPwd 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   900
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   3855
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   270
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Desconecta ODBC"
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
      Index           =   3
      Left            =   5295
      TabIndex        =   3
      ToolTipText     =   "Desconectar origen de datos ODBC"
      Top             =   1425
      Width           =   1515
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Salir"
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
      Index           =   4
      Left            =   5295
      TabIndex        =   4
      ToolTipText     =   "Salir de la pantalla"
      Top             =   1905
      Width           =   1515
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Conectar ODBC"
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
      Left            =   5295
      TabIndex        =   2
      ToolTipText     =   "Conectar a origen de datos ODBC"
      Top             =   945
      Width           =   1515
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   5745
      Top             =   5370
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
            Picture         =   "Conexiones.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conexiones.frx":40E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conexiones.frx":4242
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Archivo de Datos MDB"
      Height          =   195
      Left            =   3345
      TabIndex        =   8
      Top             =   6015
      Width           =   1635
   End
End
Attribute VB_Name = "frmConexiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Sub AgregaArchivoMDB()

    Dim Archivo As String
    Dim cc As New GCommonDialog
    Dim Glosa As String
    Dim LlaveMdb As String
    
    Glosa = "Archivos de bases de datos access (*.MDB)|*.MDB|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , App.Path, "Abrir archivo ...", "MDB", Me.hwnd)) Then
       Exit Sub
    End If
            
    If Archivo <> "" Then
        If InStr(Archivo, Chr$(0)) Then
            Archivo = Left$(Archivo, Len(Archivo) - 1)
        End If
        On Local Error Resume Next
        LlaveMdb = LlaveArchivoMdb(Archivo)
        treeMdb.Nodes.Add("EJ1", tvwChild, LlaveMdb, Archivo, 3).EnsureVisible
        treeMdb.Nodes(LlaveMdb).Key = LlaveMdb
        ReDim Preserve aArchivosMdb(UBound(aArchivosMdb) + 1)
        aArchivosMdb(UBound(aArchivosMdb)) = Archivo
        Err = 0
    End If
            
    Set cc = Nothing
    
End Sub

Private Sub CargaConexionesTreeViewODBC()

    Dim j As Integer
    Dim K As Integer
    
    Call GetDSNsAndDrivers
    
    For j = 1 To treConexiones.Nodes.Count
        For K = 1 To UBound(cState)
            If cState(K).Conexion = treConexiones.Nodes(j).Text Then
                If Not cState(K).Deleted And cState(K).tipo = TIPO_ODBC Then
                    treConexiones.Nodes(j).Tag = "Conectado"
                    treConexiones.Nodes(j).Image = 2
                End If
            End If
        Next K
    Next j
    
End Sub
Private Sub CargaConexionesTreeViewMDB()

    Dim K As Integer
    Dim Llave As String
    Dim j As Integer
    
    treeMdb.Nodes.Clear
    treeMdb.Nodes.Add(, , "EJ1", "Conexiones", 1).EnsureVisible
        
    For K = 1 To UBound(aArchivosMdb)
        Llave = LlaveArchivoMdb(aArchivosMdb(K))
        treeMdb.Nodes.Add("EJ1", tvwChild, Llave, aArchivosMdb(K), 3).EnsureVisible
        treeMdb.Nodes(Llave).Key = Llave
    Next K
    
    For j = 1 To UBound(aArchivosMdb)
        For K = 1 To UBound(cState)
            If cState(K).tipo = TIPO_MDB And cState(K).ArchivoMdb = aArchivosMdb(j) Then
                If cState(K).Conectado Then
                    treeMdb.Nodes(cState(K).LlaveMdb).Tag = "Conectado"
                    treeMdb.Nodes(cState(K).LlaveMdb).Image = 2
                End If
            End If
        Next K
    Next j
    
End Sub

Private Function LlaveArchivoMdb(ByVal Archivo As String) As String

    Dim K As Integer
    
    Dim ret As String
    
    ret = ""
    
    For K = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, K, 1) = "\" Then
            ret = Mid$(Archivo, K + 1)
            Exit For
        End If
    Next K
    
    LlaveArchivoMdb = ret
    
End Function

Private Sub cmdAccion_Click(Index As Integer)

    Dim Conexion As String
    Dim Msg As String
    Dim K As Integer
    Dim Usuario As String
    Dim Pwd As String
    Dim ArchivoMdb As String
    Dim LlaveMdb As String
        
    Dim j As Integer
    
    Select Case Index
        Case 0  'Conectar ODBC
            
            If treConexiones.SelectedItem Is Nothing Then Exit Sub
            
            If treConexiones.Nodes.Count = 0 Then Exit Sub
                                    
            Conexion = treConexiones.SelectedItem.Text
            If Conexion = "Conexiones" Then Exit Sub
                
            If treConexiones.SelectedItem.Tag = "Conectado" Then Exit Sub
            
            Call Hourglass(hwnd, True)
            
            Usuario = txtUsuario.Text
            Pwd = txtPwd.Text
            
            Call frmMain.HabilitaBotones(False)
            Call frmMain.HabilitaMenues(False)
            
            Me.Hide
            If NewConection(Conexion, TIPO_ODBC, Usuario, Pwd) Then
                
                treConexiones.SelectedItem.Tag = "Conectado"
                treConexiones.SelectedItem.Image = 2
                
                Conectado = True
                
                'frmMain.Toolbar.Buttons("cmdConectar").Enabled = False
                frmMain.Toolbar.Buttons("cmdDesconectar").Enabled = True
                frmMain.Toolbar.Buttons("cmdIniTrx").Enabled = True
                frmMain.Toolbar.Buttons("cmdBuscar").Enabled = True
                
                frmMain.Toolbar.Buttons("cmdAscendente").Enabled = True
                frmMain.Toolbar.Buttons("cmdDescendente").Enabled = True
                
                frmMain.mnuQuery.Enabled = True
                
                Call CargaConexiones
                
            End If
            
            Call frmMain.HabilitaBotones(True)
            Call frmMain.HabilitaMenues(True)
            Call frmMain.HabiBotones2
            
            Me.Show
            
            Call Hourglass(hwnd, False)
            
        Case 1  'Agregar Origen de datos
            Call WinExec("rundll32.exe shell32.dll,Control_RunDLL ODBCCP32.cpl,,1", 1)
        Case 2  'Agrega MDB
            Call AgregaArchivoMDB
        Case 3  'Desconectar
            If treConexiones.SelectedItem Is Nothing Then Exit Sub
            
            If treConexiones.Nodes.Count = 0 Then Exit Sub
                                    
            Conexion = treConexiones.SelectedItem.Text
            If Conexion = "Conexiones" Then Exit Sub
            
            If treConexiones.SelectedItem.Tag = "Conectado" Then
                Msg = "Confirma desconexión."
                If Confirma(Msg) = vbNo Then Exit Sub
                
                For K = 1 To UBound(cState)
                    If Conexion = cState(K).Conexion And cState(K).tipo = TIPO_ODBC Then
                        If cState(K).Trx = True Then
                            Msg = "Esta conexión tiene una transacción activa."
                            Msg = Msg & vbNewLine & vbNewLine & "Finalizarla."
                            
                            If Confirma(Msg) = vbYes Then
                                DBConnection(K).CommitTrans
                            Else
                                DBConnection(K).RollbackTrans
                            End If
                        End If
                        
                        cState(K).Deleted = True
                        If DBConnection(K).State > 0 Then
                            DBConnection(K).Close
                        End If
                        treConexiones.SelectedItem.Image = 3
                        treConexiones.SelectedItem.Tag = ""
                        Exit For
                    End If
                Next K
                
                Call CargaConexiones
                
                For K = 1 To UBound(cState)
                    If cState(K).Deleted = False Then
                        Exit Sub
                    End If
                Next K
                
                Conectado = False
                frmMain.Toolbar.Buttons("cmdConectar").Enabled = True
                frmMain.Toolbar.Buttons("cmdDesconectar").Enabled = False
                frmMain.mnuQuery.Enabled = False
                
                frmMain.Toolbar.Buttons("cmdIniTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdFinTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdRollTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdAscendente").Enabled = False
                frmMain.Toolbar.Buttons("cmdDescendente").Enabled = False
            End If
            
        Case 4
            Unload Me
        Case 5  'Desconecta MDB
            If treeMdb.SelectedItem Is Nothing Then Exit Sub
            
            If treeMdb.Nodes.Count = 0 Then Exit Sub
                                    
            Conexion = treeMdb.SelectedItem.Text
            
            If Conexion = "Conexiones" Then Exit Sub
            
            If treeMdb.SelectedItem.Tag = "Conectado" Then
                Msg = "Confirma desconexión."
                If Confirma(Msg) = vbNo Then Exit Sub
                
                For K = 1 To UBound(cState)
                    If Conexion = cState(K).ArchivoMdb And cState(K).tipo = TIPO_MDB Then
                        If cState(K).Trx = True Then
                            Msg = "Esta conexión tiene una transacción activa."
                            Msg = Msg & vbNewLine & vbNewLine & "Finalizarla."
                            
                            If Confirma(Msg) = vbYes Then
                                DBConnection(K).CommitTrans
                            Else
                                DBConnection(K).RollbackTrans
                            End If
                        End If
                        
                        cState(K).Deleted = True
                        If DBConnection(K).State > 0 Then
                            DBConnection(K).Close
                        End If
                        cState(K).Conectado = False
                        treeMdb.SelectedItem.Image = 3
                        treeMdb.SelectedItem.Tag = ""
                        Exit For
                    End If
                Next K
                
                Call CargaConexiones
                
                For K = 1 To UBound(cState)
                    If cState(K).Deleted = False Then
                        Exit Sub
                    End If
                Next K
                
                Conectado = False
                frmMain.Toolbar.Buttons("cmdConectar").Enabled = True
                frmMain.Toolbar.Buttons("cmdDesconectar").Enabled = False
                frmMain.mnuQuery.Enabled = False
                
                frmMain.Toolbar.Buttons("cmdIniTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdFinTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdRollTrx").Enabled = False
                frmMain.Toolbar.Buttons("cmdAscendente").Enabled = False
                frmMain.Toolbar.Buttons("cmdDescendente").Enabled = False
            End If
        Case 6  'Conecta MDB
            If treeMdb.SelectedItem Is Nothing Then Exit Sub
            
            If treeMdb.Nodes.Count = 0 Then Exit Sub
                                    
            Conexion = treeMdb.SelectedItem.Text
            If Conexion = "Conexiones" Then Exit Sub
                
            If treeMdb.SelectedItem.Tag = "Conectado" Then Exit Sub
            
            Call Hourglass(hwnd, True)
            Call frmMain.HabilitaBotones(False)
            Call frmMain.HabilitaMenues(False)
            
            Usuario = txtUsuario.Text
            Pwd = txtPwd.Text
            
            ArchivoMdb = Conexion
            LlaveMdb = treeMdb.SelectedItem.Key
            
            If opt97.Value <> False Then
                Conexion = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                    & "Persist Security Info=False;Data Source=" & Conexion
                        
            Else
                Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                    & "Persist Security Info=False;Data Source=" & Conexion
            End If
            
            Me.Hide
            If NewConection(Conexion, TIPO_MDB, Usuario, Pwd, ArchivoMdb, LlaveMdb) Then
                
                treeMdb.SelectedItem.Tag = "Conectado"
                treeMdb.SelectedItem.Image = 2
                
                Conectado = True
                
                frmMain.Toolbar.Buttons("cmdDesconectar").Enabled = True
                frmMain.Toolbar.Buttons("cmdIniTrx").Enabled = True
                frmMain.Toolbar.Buttons("cmdBuscar").Enabled = True
                
                frmMain.Toolbar.Buttons("cmdAscendente").Enabled = True
                frmMain.Toolbar.Buttons("cmdDescendente").Enabled = True
                
                frmMain.mnuQuery.Enabled = True
                
                Call CargaConexiones
            End If
            
            Call frmMain.HabilitaBotones(True)
            Call frmMain.HabilitaMenues(True)
            Call frmMain.HabiBotones2
            
            Me.Show
            Call Hourglass(hwnd, False)
        Case 7  'REMOVER MDB
            If treeMdb.SelectedItem Is Nothing Then Exit Sub
            
            If treeMdb.Nodes.Count = 0 Then Exit Sub
                                    
            Conexion = treeMdb.SelectedItem.Text
            If Conexion = "Conexiones" Then Exit Sub
                
            If treeMdb.SelectedItem.Tag = "Conectado" Then Exit Sub
            
            Msg = "Confirma remover base .mdb de lista."
                        
            If Confirma(Msg) = vbYes Then
                treeMdb.Nodes.Remove treeMdb.SelectedItem.Key
                ReDim aArchivosMdb(0)
                For j = 2 To treeMdb.Nodes.Count
                    ReDim Preserve aArchivosMdb(j - 1)
                    aArchivosMdb(j - 1) = treeMdb.Nodes(j).Text
                Next j
            End If
    End Select
    
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    Call CargaConexionesTreeViewODBC
    Call CargaConexionesTreeViewMDB
        
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, "Conectar a origen de datos")
    
    picDraw.Refresh
    
End Sub


Sub GetDSNsAndDrivers()
  
    On Error Resume Next
    
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long     'handle to the environment

    treConexiones.Nodes.Clear
    treConexiones.Nodes.Add(, , "EJ1", "Conexiones", 1).EnsureVisible
    
    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            
            sDSN = VBA.Left(sDSNItem, iDSNLen)
            sDRV = VBA.Left(sDRVItem, iDRVLen)
            
            If sDSN <> Space(iDSNLen) Then
                treConexiones.Nodes.Add("EJ1", tvwChild, sDSN, sDSN, 3).EnsureVisible
                treConexiones.Nodes(sDSN).Key = sDSN
                'cboDSNList.AddItem sDSN
                'cboDrivers.AddItem sDRV
            End If
        Loop
    End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmConexiones = Nothing
    
End Sub


Private Sub treConexiones_DblClick()
    Call cmdAccion_Click(0)
End Sub

Private Sub treeMdb_DblClick()
    Call cmdAccion_Click(6)
End Sub


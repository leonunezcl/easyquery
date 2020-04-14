VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTablas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browser de Tablas"
   ClientHeight    =   5595
   ClientLeft      =   1815
   ClientTop       =   1965
   ClientWidth     =   4800
   Icon            =   "BrowserTablas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4230
      Top             =   990
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
            Picture         =   "BrowserTablas.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowserTablas.frx":0C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowserTablas.frx":0D7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCargar"
            Object.ToolTipText     =   "Cargar tablas de conexión"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEstructura"
            Object.ToolTipText     =   "Ver Estructura de la tabla"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdExpandir"
            ImageIndex      =   3
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFiltro 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   840
      Width           =   3075
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2250
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowserTablas.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowserTablas.frx":1976
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treTablas 
      Height          =   4005
      Left            =   60
      TabIndex        =   0
      Top             =   1530
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   7064
      _Version        =   393217
      Indentation     =   229
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageCombo imgConexiones 
      Height          =   330
      Left            =   1620
      TabIndex        =   2
      ToolTipText     =   "Conexiones activas"
      Top             =   480
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgLst"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Filtrar tablas según"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   870
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione conexión"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione tabla :"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1290
      Width           =   1275
   End
End
Attribute VB_Name = "frmTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaConexiones()

    Dim i As Integer
    Dim K As Integer
    Dim ci As ComboItem
    
    Dim ArrayCount As Integer
    Dim ArrayDoc As Integer

    ArrayCount = UBound(DBConnection)
    ArrayDoc = UBound(Document)
    
    For K = 1 To ArrayDoc
        If Not fState(K).Deleted Then
            imgConexiones.ComboItems.Clear
            
            For i = 1 To ArrayCount
                If Not cState(i).Deleted Then
                    
                    If cState(i).tipo = TIPO_ODBC Then
                        Set ci = imgConexiones.ComboItems.Add(1, cState(i).Conexion, _
                        cState(i).Conexion, 2, 2, 0)
                    Else
                        Set ci = imgConexiones.ComboItems.Add(1, cState(i).ArchivoMdb, _
                        cState(i).LlaveMdb, 2, 2, 0)
                    End If
                    
                    ci.Indentation = 2
                    
                End If
            Next i
        End If
    Next K
    
    Set ci = Nothing
    
End Sub


Private Sub CargarTablas()

    On Local Error GoTo ErrorCargaTablas
    
    Dim Tabla As String
    Dim ci As ComboItem
    Dim ActiveConexion As Integer
    Dim Rs As New ADODB.Recordset
    Dim RsNombre As New ADODB.Recordset
    Dim Nombre As String
    
    Dim sql As String
    Dim i As Integer
    Dim Filtro As String
    
    If imgConexiones.SelectedItem Is Nothing Then
        MsgBox "Seleccione una conexión.", vbCritical
        Exit Sub
    End If
    
    Set ci = imgConexiones.SelectedItem
    
    ActiveConexion = ConexionActiva(ci.Text)
    
    Filtro = Trim$(txtFiltro.Text)
    
    Call Hourglass(hwnd, True)
    
    Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    
    treTablas.Nodes.Clear
    treTablas.Nodes.Add(, , "EJ1", "Tablas", 1).EnsureVisible
    
    Do While Not Rs.EOF
        If Rs(3) = "TABLE" And Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
            ValidateRect treTablas.hwnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect treTablas.hwnd, 0&, 0&
            
            If Len(Filtro) = 0 Then
                'If Not IsNull(Rs!Description) Then
                '    Tabla = Trim$(Rs!table_name) & " - " & Trim$(Rs!Description)
                'Else
                    Tabla = Trim$(Rs!table_name)
                'end If
                
                Nombre = Trim$(Tabla)
                            
                treTablas.Nodes.Add("EJ1", tvwChild, "T" & i, Nombre, 2).EnsureVisible
                
                treTablas.Nodes("T" & i).Tag = Tabla
                
                i = i + 1
            Else
                'If Not IsNull(Rs!Description) Then
                '    Tabla = Trim$(Rs!table_name) & " - " & Trim$(Rs!Description)
                'Else
                    Tabla = Trim$(Rs!table_name)
                'End If
                
                Nombre = Trim$(Tabla)
                            
                If Nombre Like Filtro Then
                    treTablas.Nodes.Add("EJ1", tvwChild, "T" & i, Nombre, 2).EnsureVisible
                
                    treTablas.Nodes("T" & i).Tag = Tabla
                
                    i = i + 1
                End If
            End If
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    InvalidateRect treTablas.hwnd, 0&, 0&
    
    Set ci = Nothing
    Set Rs = Nothing
    Set RsNombre = Nothing
    
    GoTo SalirCargaTablas
    
ErrorCargaTablas:
    MsgBox "CargaTablas : " & Err & " " & Error$, vbCritical
    Resume SalirCargaTablas
    
SalirCargaTablas:
    Err = 0
    Call Hourglass(hwnd, False)
    
End Sub



Private Sub Form_Load()

    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    
    treTablas.Nodes.Clear
    treTablas.Nodes.Add(, , "EJ1", "Tablas", 1).EnsureVisible
    
    Call CargaConexiones
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmTablas = Nothing
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim Nodo As Node
    Dim ci As ComboItem
    
    If Button.Index = 1 Then
        Call CargarTablas
    ElseIf Button.Index = 2 Then
                
        If imgConexiones.SelectedItem Is Nothing Then
            MsgBox "Seleccione una conexión.", vbCritical
            Exit Sub
        End If
    
        Set ci = imgConexiones.SelectedItem
        
        If treTablas.SelectedItem Is Nothing Then
            MsgBox "Seleccione una tabla.", vbCritical
            Exit Sub
        End If
        
        If treTablas.SelectedItem.Text <> "Tablas" Then
            Me.Hide
            frmVerTabla.TABLA_ORIGEN = treTablas.SelectedItem.Tag
            frmVerTabla.CONEXION_ORIGEN = ci.Text
            frmVerTabla.Show vbModal
            Me.Show
        Else
            MsgBox "Debe seleccionar una tabla.", vbCritical
        End If
    ElseIf Button.Index = 4 Then
        If Button.Value = tbrPressed Then
            Me.Height = Toolbar1.Height + 350
        Else
            Me.Height = 5925
        End If
    End If
    
    Set ci = Nothing
    Set Nodo = Nothing
    
End Sub



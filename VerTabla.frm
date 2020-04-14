VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmVerTabla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de Estructura de Tabla"
   ClientHeight    =   4110
   ClientLeft      =   2415
   ClientTop       =   4620
   ClientWidth     =   9480
   Icon            =   "VerTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3210
      Top             =   2640
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
            Picture         =   "VerTabla.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VerTabla.frx":0C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAceptar"
            Object.ToolTipText     =   "Ver estructura"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdImprimir"
            Object.ToolTipText     =   "Imprimir Estructura"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTabla 
      Height          =   285
      Left            =   4425
      TabIndex        =   2
      Top             =   765
      Width           =   4950
   End
   Begin FPSpreadADO.fpSpread fpTabla 
      Height          =   2985
      Left            =   135
      TabIndex        =   0
      Top             =   1095
      Width           =   9225
      _Version        =   196608
      _ExtentX        =   16272
      _ExtentY        =   5265
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      ScrollBars      =   2
      SpreadDesigner  =   "VerTabla.frx":33D2
      UserResize      =   0
   End
   Begin MSComctlLib.ImageCombo imgConexiones 
      Height          =   330
      Left            =   135
      TabIndex        =   3
      ToolTipText     =   "Conexiones activas"
      Top             =   735
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgLst"
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   3555
      Top             =   570
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
            Picture         =   "VerTabla.frx":381D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VerTabla.frx":7039
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VerTabla.frx":7195
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione conexión"
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
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   1785
   End
   Begin VB.Label lblTabla 
      AutoSize        =   -1  'True
      Caption         =   "Digite nombre de la tabla"
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
      Left            =   4425
      TabIndex        =   1
      Top             =   570
      Width           =   2145
   End
End
Attribute VB_Name = "frmVerTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TABLA_ORIGEN As String
Public CONEXION_ORIGEN As String
Public VISTA As String
Private Rs As New ADODB.Recordset
Private Campo As Field
Private Sub CargaInfoTabla(ByVal Tabla As String, ByVal ActiveConexion As Integer)

    Dim f As Integer
    f = 1
    Dim RsField As New ADODB.Recordset
    
    Do While Not Rs.EOF
        If Rs(3) = "TABLE" And TABLA_ORIGEN <> "" Then
            If UCase(Left(Rs!table_name, 4)) <> "MSYS" Then
                If UCase(Left(Rs!table_name, 11)) <> "SWITCHBOARD" Then
                    If Rs!table_name = Tabla Then
                        Set RsField = DBConnection(ActiveConexion).OpenSchema(adSchemaColumns)
                        
                        Do While Not RsField.EOF
                            If RsField(2) = Rs(2) Then
                                For Each Campo In RsField.Fields
                                    If Campo.Name = "COLUMN_NAME" Then
                                        Call fpTabla.SetText(1, f, Campo.Value)
                                    ElseIf Campo.Name = "DESCRIPTION" Then
                                        Call fpTabla.SetText(2, f, IIf(IsNull(Campo.Value), "", Campo.Value))
                                    ElseIf Campo.Name = "DATA_TYPE" Then
                                        Call fpTabla.SetText(3, f, TipoDeCampo(Campo.Value))
                                    ElseIf Campo.Name = "CHARACTER_MAXIMUM_LENGTH" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "DATETIME_PRECISION" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "NUMERIC_PRECISION" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "NUMERIC_SCALE" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(5, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "IS_NULLABLE" Then
                                        If (Campo.Attributes And adFldIsNullable) Then
                                            Call fpTabla.SetText(6, f, "Si")
                                        Else
                                            Call fpTabla.SetText(6, f, "No")
                                        End If
                                    End If
                                Next
                                f = f + 1
                            End If
                            RsField.MoveNext
                        Loop
                        RsField.Close
                        Exit Do
                    End If
                End If
            End If
        ElseIf Rs(3) = "VIEW" And VISTA <> "" Then
            If UCase(Left(Rs!table_name, 4)) <> "MSYS" Then
                If UCase(Left(Rs!table_name, 11)) <> "SWITCHBOARD" Then
                    If Rs!table_name = Tabla Then
                        Set RsField = DBConnection(ActiveConexion).OpenSchema(adSchemaColumns)
                        
                        Do While Not RsField.EOF
                            If RsField(2) = Rs(2) Then
                                For Each Campo In RsField.Fields
                                    If Campo.Name = "COLUMN_NAME" Then
                                        Call fpTabla.SetText(1, f, Campo.Value)
                                    ElseIf Campo.Name = "DESCRIPTION" Then
                                        Call fpTabla.SetText(2, f, IIf(IsNull(Campo.Value), "", Campo.Value))
                                    ElseIf Campo.Name = "DATA_TYPE" Then
                                        Call fpTabla.SetText(3, f, TipoDeCampo(Campo.Value))
                                    ElseIf Campo.Name = "CHARACTER_MAXIMUM_LENGTH" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "DATETIME_PRECISION" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "NUMERIC_PRECISION" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(4, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "NUMERIC_SCALE" Then
                                        If Not IsNull(Campo.Value) Then
                                            Call fpTabla.SetText(5, f, Campo.Value)
                                        End If
                                    ElseIf Campo.Name = "IS_NULLABLE" Then
                                        If (Campo.Attributes And adFldIsNullable) Then
                                            Call fpTabla.SetText(6, f, "Si")
                                        Else
                                            Call fpTabla.SetText(6, f, "No")
                                        End If
                                    End If
                                Next
                                f = f + 1
                            End If
                            RsField.MoveNext
                        Loop
                        RsField.Close
                        Exit Do
                    End If
                End If
            End If
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
    
End Sub

Private Sub VerEstructuraTabla()

    On Local Error GoTo ErrorVerEstructuraTabla
    
    Dim Tabla As String
    Dim ci As ComboItem
    Dim ActiveConexion As Integer
    Dim Campo As Field
    Dim f As Integer
    Dim j As Integer
    Dim Found As Boolean
    
    If TABLA_ORIGEN <> "" Then
        txtTabla.Text = TABLA_ORIGEN
        For f = 1 To imgConexiones.ComboItems.Count
            If imgConexiones.ComboItems(f).Text = CONEXION_ORIGEN Then
                imgConexiones.ComboItems(f).Selected = True
                Exit For
            End If
        Next f
    ElseIf VISTA <> "" Then
        txtTabla.Text = VISTA
        For f = 1 To imgConexiones.ComboItems.Count
            If imgConexiones.ComboItems(f).Text = CONEXION_ORIGEN Then
                imgConexiones.ComboItems(f).Selected = True
                Exit For
            End If
        Next f
    End If
    
    Tabla = Trim$(txtTabla.Text)
    
    If CONEXION_ORIGEN = "" Then
        If imgConexiones.SelectedItem Is Nothing Then
            MsgBox "Seleccione una conexión.", vbCritical
            Exit Sub
        End If
    End If
    
    If Tabla = "" Then
        MsgBox "Ingrese nombre de la tabla.", vbCritical
        Exit Sub
    End If
    
    Set ci = imgConexiones.SelectedItem
    
    If CONEXION_ORIGEN = "" Then
        ActiveConexion = ConexionActiva(ci.Text)
    Else
        ActiveConexion = ConexionActiva(CONEXION_ORIGEN)
    End If
    
    fpTabla.MaxRows = 0
    fpTabla.MaxRows = 100
    
    If VISTA = "" Then
        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    Else
        Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "VIEW"))
    End If
                
    Call Hourglass(hwnd, True)
            
    Call CargaInfoTabla(Tabla, ActiveConexion)
    
    f = fpTabla.DataRowCnt + 2
    fpTabla.Col = 1
    fpTabla.Row = f
    fpTabla.Col2 = 1
    fpTabla.Row2 = f
    fpTabla.FontBold = True
    Call fpTabla.SetText(1, f, "INDICES")
    
    f = fpTabla.DataRowCnt + 1
    
    Set Rs = DBConnection(ActiveConexion).OpenSchema(adSchemaIndexes)
        
    If f >= fpTabla.MaxRows Then
        fpTabla.MaxRows = fpTabla.MaxRows + 1
    End If
            
    Do While Not Rs.EOF
        If Rs!table_name = Tabla Then
            For Each Campo In Rs.Fields
                If Campo.Name = "INDEX_NAME" Then
                    fpTabla.Col = 1
                    fpTabla.Row = f
                    fpTabla.FontBold = True
                    Call fpTabla.SetText(1, f, "Nombre Indice")
                    
                    fpTabla.Col = 2
                    fpTabla.Row = f
                    fpTabla.AllowCellOverflow = True
                    Call fpTabla.SetText(2, f, Campo.Value)
                    f = f + 1
                    Found = True
                ElseIf Campo.Name = "PRIMARY_KEY" Then
                    fpTabla.Col = 1
                    fpTabla.Row = f
                    fpTabla.FontBold = True
                    Call fpTabla.SetText(1, f, "Principal")
                    
                    fpTabla.Col = 2
                    fpTabla.Row = f
                    fpTabla.AllowCellOverflow = True
                    
                    If Left$(Campo.Value, 1) = "F" Then
                        Call fpTabla.SetText(2, f, "No")
                    Else
                        Call fpTabla.SetText(2, f, "Si")
                    End If
                    f = f + 1
                    Found = True
                ElseIf Campo.Name = "COLUMN_NAME" Then
                    fpTabla.Col = 1
                    fpTabla.Row = f
                    fpTabla.FontBold = True
                    Call fpTabla.SetText(1, f, "Campo(s)")
                    
                    fpTabla.Col = 2
                    fpTabla.Row = f
                    fpTabla.AllowCellOverflow = True
                    Call fpTabla.SetText(2, f, Campo.Value)
                    f = f + 2
                    Found = True
                End If
                
                If Found Then
                    If f >= fpTabla.MaxRows Then
                        fpTabla.MaxRows = fpTabla.MaxRows + 2
                    End If
                End If
            Next
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
        
    GoTo SalirVerEstructuraTabla
    
ErrorVerEstructuraTabla:
    MsgBox "VerEstructuraTabla : " & Err & " " & Error$, vbCritical
    Resume SalirVerEstructuraTabla
    
SalirVerEstructuraTabla:
    If fpTabla.DataRowCnt <= 10 Then
       fpTabla.MaxRows = 10
    Else
        fpTabla.MaxRows = fpTabla.DataRowCnt
    End If
    
    Call Hourglass(hwnd, False)
    
    Set ci = Nothing
    Set Rs = Nothing
    
    If fpTabla.DataRowCnt < 10 Then fpTabla.MaxRows = fpTabla.MaxRows + (10 - fpTabla.DataRowCnt)
    Err = 0
    Call Hourglass(hwnd, False)

End Sub
Private Sub Form_Load()

    CenterWindow hwnd
    
    Call CargaConexiones
    
    If TABLA_ORIGEN <> "" Then
        Me.Caption = "Ver estructura de Tabla"
    Else
        Me.Caption = "Ver estructura de Vista"
    End If
    
    If CONEXION_ORIGEN <> "" Then Call VerEstructuraTabla
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmVerTabla = Nothing
    
End Sub

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
                    
                    If cState(i).Tipo = TIPO_ODBC Then
                        Set ci = imgConexiones.ComboItems.Add(1, cState(i).Conexion, _
                        cState(i).Conexion, 1, 1, 0)
                    Else
                        Set ci = imgConexiones.ComboItems.Add(1, cState(i).ArchivoMdb, _
                        cState(i).LlaveMdb, 1, 1, 0)
                    End If
                    
                    ci.Indentation = 2
                    
                End If
            Next i
        End If
    Next K
    
    Set ci = Nothing
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 1 Then
        Call VerEstructuraTabla
    ElseIf Button.Index = 3 Then
        frmPreview.lHwnd = fpTabla.hwnd
        frmPreview.Estructura = True
        frmPreview.Show vbModal
    End If
    
End Sub


Private Sub txtTabla_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdAceptar"))
    End If
    
End Sub


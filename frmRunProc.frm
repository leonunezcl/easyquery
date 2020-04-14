VERSION 5.00
Begin VB.Form frmRunProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutar procedimiento"
   ClientHeight    =   4065
   ClientLeft      =   1545
   ClientTop       =   2745
   ClientWidth     =   7605
   ControlBox      =   0   'False
   Icon            =   "frmRunProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInfo 
      Height          =   300
      Left            =   2175
      TabIndex        =   11
      Text            =   "CALL "
      Top             =   360
      Width           =   5385
   End
   Begin VB.CommandButton cmd 
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
      Height          =   375
      Index           =   1
      Left            =   4125
      TabIndex        =   4
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ejecutar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1905
      TabIndex        =   3
      Top             =   3585
      Width           =   1215
   End
   Begin VB.Frame fraPar 
      Caption         =   "Parámetros procedimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   90
      TabIndex        =   1
      Top             =   705
      Width           =   7470
      Begin VB.PictureBox picMain 
         BorderStyle     =   0  'None
         Height          =   2340
         Left            =   75
         ScaleHeight     =   2340
         ScaleWidth      =   7305
         TabIndex        =   5
         Top             =   240
         Width           =   7305
         Begin VB.VScrollBar VScroll1 
            Height          =   2355
            Left            =   7110
            TabIndex        =   9
            Top             =   -15
            Width           =   180
         End
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   0
            ScaleHeight     =   330
            ScaleWidth      =   7080
            TabIndex        =   6
            Top             =   -15
            Width           =   7080
            Begin VB.TextBox txtValor 
               Height          =   285
               Index           =   0
               Left            =   2775
               TabIndex        =   7
               Top             =   15
               Width           =   4245
            End
            Begin VB.Label lblParam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PARAM"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   8
               Top             =   0
               Width           =   570
            End
         End
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Información anexa"
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
      Left            =   90
      TabIndex        =   10
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label lblProc 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   2190
      TabIndex        =   2
      Top             =   90
      Width           =   5370
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "Nombre procedimiento : "
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2085
   End
End
Attribute VB_Name = "frmRunProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private total As Integer
Private ActiveConexion As Integer
Private cat As New ADOX.Catalog
Private Rs As New ADODB.Recordset
Private cmdx As New ADODB.Command
Private prm As New ADODB.Parameter
Private Campo As ADODB.Field
Public TABLA_ORIGEN As String
Public CONEXION_ORIGEN As String
'cargar parametros de base ndb
Private Sub CargaParametrosMDB(ByVal Indice As Integer)

    On Local Error GoTo ErrorCargaParametrosMDB
    
    Dim K As Integer
    Dim j As Integer
    
    Dim buffer As String
    Dim Biblio As String
    Dim largo As Long
    Dim tipo As String
    
    Set cat.ActiveConnection = DBConnection(Indice)
    Set cmdx = cat.Procedures(TABLA_ORIGEN).Command
    
    K = 0
    'ciclar x todos los parametros
    For Each prm In cmdx.Parameters
        If K = 0 Then
            Call NombreParametro(prm.Name, K)
            lblParam(K).Caption = lblParam(K).Caption & " (" & TipoDeCampo(prm.Type) & ")"
            lblParam(K).AutoSize = True
            K = K + 1
        Else
            Load lblParam(K)
            lblParam(K).Left = lblParam(0).Left
            lblParam(K).Height = lblParam(0).Height
            lblParam(K).Top = lblParam(K - 1).Top + 300
            lblParam(K).Visible = True
            Call NombreParametro(prm.Name, K)
            picInfo.Height = picInfo.Height + 300
            
            lblParam(K).Caption = lblParam(K).Caption & " (" & TipoDeCampo(prm.Type) & ")"
            lblParam(K).AutoSize = True
            
            Load txtValor(K)
            txtValor(K).Left = txtValor(0).Left
            txtValor(K).Height = txtValor(0).Height
            txtValor(K).Top = txtValor(K - 1).Top + 300
            txtValor(K).Visible = True
            
            K = K + 1
        End If
    Next
    
    txtInfo.Text = lblProc.Caption
        
    total = K - 1
        
    Exit Sub
    
ErrorCargaParametrosMDB:
    MsgBox "CargaParametrosMDB : " & Err & " " & Error$
    Resume SalirCargaParametrosMDB
    
SalirCargaParametrosMDB:
    Err = 0

End Sub

'cargar parametros desde odbc
Private Sub CargaParametrosODBC(ByVal Indice As Integer)

    On Local Error GoTo ErrorCargaParametrosODBC
    
    Dim K As Integer
    Dim j As Integer
    
    Dim buffer As String
    Dim Biblio As String
    Dim largo As Long
    Dim tipo As String
    
    Set Rs = DBConnection(Indice).OpenSchema(adSchemaProcedures)
    
    K = 0
    buffer = ""
    Do While Not Rs.EOF
        If Rs!procedure_name = TABLA_ORIGEN Then
            Rs.Close
            Set Rs = DBConnection(Indice).OpenSchema(adSchemaProcedureParameters)
            
            Do While Not Rs.EOF
                If Rs!procedure_name = TABLA_ORIGEN Then
                    For Each Campo In Rs.Fields
                        If Campo.Name = "PROCEDURE_SCHEMA" Then
                            If Not IsNull(Campo.Value) Then
                                Biblio = Campo.Value
                            End If
                        ElseIf Campo.Name = "CHARACTER_MAXIMUM_LENGTH" Then
                            If Not IsNull(Campo.Value) Then
                                largo = Campo.Value
                                tipo = "C"
                            End If
                        ElseIf Campo.Name = "NUMERIC_PRECISION" Then
                            If Not IsNull(Campo.Value) Then
                                largo = Campo.Value
                                tipo = "N"
                            End If
                        ElseIf Campo.Name = "PARAMETER_NAME" Then
                            If K = 0 Then
                                Call NombreParametro(Campo.Value, K)
                            Else
                                Load lblParam(K)
                                lblParam(K).Left = lblParam(0).Left
                                lblParam(K).Height = lblParam(0).Height
                                lblParam(K).Top = lblParam(K - 1).Top + 300
                                lblParam(K).Visible = True
                                Call NombreParametro(Campo.Value, K)
                                picInfo.Height = picInfo.Height + 300
                            End If
                        ElseIf Campo.Name = "TYPE_NAME" Then
                            If K = 0 Then
                                lblParam(K).Caption = lblParam(K).Caption & " (" & Campo.Value & ")"
                                lblParam(K).AutoSize = True
                                K = K + 1
                            Else
                                lblParam(K).Caption = lblParam(K).Caption & " (" & Campo.Value & ")"
                                lblParam(K).AutoSize = True
                                
                                Load txtValor(K)
                                txtValor(K).Left = txtValor(0).Left
                                txtValor(K).Height = txtValor(0).Height
                                txtValor(K).Top = txtValor(K - 1).Top + 300
                                txtValor(K).Visible = True
                                
                                K = K + 1
                            End If
                        End If
                    Next
                    txtValor(K - 1).MaxLength = largo
                    txtValor(K - 1).Tag = tipo
                End If
                Rs.MoveNext
            Loop
            Exit Do
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
                    
    txtInfo.Text = "CALL " & Biblio & "." & lblProc.Caption
        
    total = K - 1
        
    Exit Sub
    
ErrorCargaParametrosODBC:
    MsgBox "CargaParametrosODBC : " & Err & " " & Error$
    Resume SalirCargaParametrosODBC
    
SalirCargaParametrosODBC:
    Err = 0
    
End Sub
Private Sub CargarParametros()
    
    
        
    ActiveConexion = ConexionActiva(CONEXION_ORIGEN)
        
    If cState(ActiveConexion).tipo = TIPO_ODBC Then
        Call CargaParametrosODBC(ActiveConexion)
    Else
        Call CargaParametrosMDB(ActiveConexion)
    End If
        
End Sub

Private Function EjecutaProcedimiento() As Boolean

    On Local Error GoTo ErrorEjecutaProcedimiento

    Dim ret As Boolean
    Dim K As Integer
    Dim sql As String
    Dim Msg As String
    
    ret = True
    
    Call Hourglass(hwnd, True)
        
    If cState(ActiveConexion).tipo = TIPO_ODBC Then
        sql = txtInfo.Text & " ("
        
        For K = 0 To total
            If txtValor(K).Tag = "C" Then
                sql = sql & "'" & txtValor(K).Text & "' , "
            Else
                sql = sql & " " & txtValor(K).Text & " , "
            End If
        Next K
                
        sql = Left$(sql, Len(sql) - 3) & ")"
    Else
        K = 0
        For Each prm In cmdx.Parameters
            prm.Value = txtValor(K).Text
            K = K + 1
        Next
    End If
    
    Msg = "Confirma ejecutar : " & vbNewLine
    Msg = Msg & vbNewLine
    If cState(ActiveConexion).tipo = TIPO_ODBC Then
        Msg = Msg & sql
    Else
        Msg = Msg & lblProc.Caption
    End If
    
    If Confirma(Msg) = vbYes Then
        If cState(ActiveConexion).tipo = TIPO_ODBC Then
            frmMain.ActiveForm.txtQuery.Text = sql
        Else
            frmMain.ActiveForm.txtQuery.Text = cmdx.CommandText
        End If
        frmMain.ActiveForm.FormateaSentencias
    
        Me.Hide
        
        If cState(ActiveConexion).tipo = TIPO_ODBC Then
            Call frmMain.EjecutaQuery2
        Else
            Call frmMain.EjecutaQuery2(cmdx)
        End If
    Else
        ret = False
    End If
    
    GoTo SalirEjecutaProcedimiento
    
ErrorEjecutaProcedimiento:
    ret = False
    MsgBox "EjecutaProcedimiento : " & Err & " " & Error$, vbCritical
    Resume SalirEjecutaProcedimiento
    
SalirEjecutaProcedimiento:
    Me.Show
    Call Hourglass(hwnd, False)
    EjecutaProcedimiento = ret
    Err = 0
    
End Function

Private Sub NombreParametro(ByVal Nombre As String, ByVal K As Integer)

    Dim j As Integer
    
    If InStr(Nombre, "!") = 0 Then
        lblParam(K).Caption = Nombre
        lblParam(K).Tag = Nombre
    Else
        For j = Len(Nombre) To 1 Step -1
            If Mid$(Nombre, j, 1) = "!" Then
                lblParam(K).Caption = Mid$(Nombre, j + 1)
                lblParam(K).Tag = Nombre
                Exit For
            End If
        Next j
    End If
            
End Sub


Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    If Index = 0 Then
        If EjecutaProcedimiento() Then
            MsgBox "Procedimiento ejecutado con éxito!", vbInformation
            Unload Me
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()

    On Local Error Resume Next
    
    txtValor(0).SetFocus
    
    Err = 0
    
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    Hourglass hwnd, True
    
    lblProc.Caption = TABLA_ORIGEN
    
    CargarParametros
    
    If (picMain.Height - picInfo.Height) > 0 Then
        VScroll1.Max = picMain.Height - picInfo.Height
    Else
        VScroll1.Max = picInfo.Height - picMain.Height
    End If

    'Set VScrollBar LargeChange and SmallChange
    VScroll1.LargeChange = VScroll1.Max \ 10
    VScroll1.SmallChange = VScroll1.Max \ 50
    
    Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
        
    Set frmRunProc = Nothing
    
End Sub


Private Sub VScroll1_Change()
    picInfo.Top = -VScroll1.Value
End Sub



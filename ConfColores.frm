VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConfColores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Colores"
   ClientHeight    =   5610
   ClientLeft      =   900
   ClientTop       =   1650
   ClientWidth     =   4935
   Icon            =   "ConfColores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   367
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   10
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   3585
      TabIndex        =   9
      ToolTipText     =   "Editar instruaccion SQL"
      Top             =   2670
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "A&plicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   3585
      TabIndex        =   8
      ToolTipText     =   "Aplicar cambios a hojas de consultas activas"
      Top             =   2190
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   3585
      TabIndex        =   7
      ToolTipText     =   "Eliminar instruccion SQL de lista"
      Top             =   1710
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   3585
      TabIndex        =   6
      ToolTipText     =   "Configurar color de instruccion SQL"
      Top             =   1230
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3585
      TabIndex        =   5
      ToolTipText     =   "Agregar instruccon SQL"
      Top             =   750
      Width           =   1245
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
      Height          =   435
      Index           =   0
      Left            =   3585
      TabIndex        =   4
      ToolTipText     =   "Salir de la pantalla"
      Top             =   270
      Width           =   1245
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2145
      Left            =   405
      TabIndex        =   2
      Top             =   3420
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3784
      _Version        =   393217
      TextRTF         =   $"ConfColores.frx":08CA
   End
   Begin VB.ListBox lis 
      Height          =   2790
      Left            =   405
      TabIndex        =   1
      Top             =   270
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Instrucción SQL"
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
      Left            =   435
      TabIndex        =   3
      Top             =   3150
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sentencias registradas"
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
      Left            =   405
      TabIndex        =   0
      Top             =   30
      Width           =   1950
   End
End
Attribute VB_Name = "frmConfColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Sub cmdAccion_Click(Index As Integer)

    Dim Color As Long
    Dim Sentencia As String
    Dim OldSentencia As String
    Dim msg As String
    Dim k As Integer
    Dim j As Integer
    Dim Sql As String
        
    Sql = rtb.Text
    
    Select Case Index
        Case 0  'Salir
            
            Unload Me
        Case 1  'Agregar
            Sentencia = InputBox("Digite sentencia a agregar", App.Title)
            If Sentencia <> "" Then
                lis.AddItem Sentencia
                rtb.Text = Sql & Sentencia & " "
                Call CargaSentencias
            End If
        Case 2  'Color
            If lis.ListIndex <> -1 Then
                Sentencia = lis.Text
                Color = ShowColor(hwnd)
                If Color <> -1 Then
                    Call ColorSQL(rtb, Sentencia, Color)
                    lis.itemData(lis.ListIndex) = Color
                    Call ColorSQL(rtb, Sentencia, Color)
                End If
            Else
                MsgBox "Seleccione sentencia.", vbInformation
            End If
        Case 3  'Eliminar
            If lis.ListIndex <> -1 Then
                Sentencia = lis.Text
                msg = "Confirma eliminar sentencia : " & Sentencia
                If Confirma(msg) = vbNo Then Exit Sub
                lis.RemoveItem lis.ListIndex
                Sentencia = ""
                For k = 0 To lis.ListCount - 1
                    Sentencia = Sentencia & lis.List(k) & " "
                Next k
                rtb.Text = Sentencia
            End If
        Case 4  'Aplicar
            
            Dim f As Form
            
            ReDim aSentencias(0)
            
            For k = 0 To lis.ListCount - 1
                Sentencia = lis.List(k)
                Color = lis.itemData(k)
                ReDim Preserve aSentencias(k + 1)
                aSentencias(k + 1).Glosa = Sentencia
                aSentencias(k + 1).Color = Color
                Call GrabaIni(C_INI, "Sentencias", "Sentencia" & k + 1, Sentencia)
                Call GrabaIni(C_INI, "Sentencias", "Color" & k + 1, Color)
            Next k
            
            Call GrabaIni(C_INI, "Sentencias", "nSentencias", CStr(UBound(aSentencias)))
            
            Dim ArrayCount As Integer

            ArrayCount = UBound(Document)
        
            For k = 1 To ArrayCount
                If Not fState(k).Deleted Then
                    Set f = Document(k)
                    For j = 1 To UBound(aSentencias)
                        Call ColorSQL(f.txtQuery, aSentencias(j).Glosa, aSentencias(j).Color)
                    Next j
                End If
            Next
    
            Set f = Nothing
            
            Unload Me
        Case 5  'Editar
            If lis.ListIndex <> -1 Then
                OldSentencia = lis.Text
                Sentencia = InputBox("Editar sentencia ", App.Title, OldSentencia)
                If Sentencia <> "" Then
                    lis.Text = Sentencia
                    rtb.Text = Sql & Sentencia
                End If
            End If
    End Select
    
End Sub


Private Sub Form_Load()

    Dim k As Integer
    Dim Sql As String
    
    CenterWindow hwnd
    
    For k = 1 To UBound(aSentencias)
        Sql = Sql & aSentencias(k).Glosa & " "
    Next k
    
    rtb.Text = Sql
    
    For k = 1 To UBound(aSentencias)
        lis.AddItem aSentencias(k).Glosa
        lis.itemData(lis.NewIndex) = aSentencias(k).Color
    Next k
    
    Call CargaSentencias
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, "Configurar SQL")
    
    picDraw.Refresh
           
End Sub


Private Sub CargaSentencias()

    Dim k As Integer
    
    Dim CHARRANGE As ECharrange
    
    Dim ret As Long
    
    ret = SendMessage(rtb.hwnd, EM_EXGETSEL, 0, CHARRANGE)
    
    For k = 1 To UBound(aSentencias)
        Call ColorSQL(rtb, aSentencias(k).Glosa, aSentencias(k).Color)
    Next k
    
    rtb.SelColor = glbColorSql
    rtb.SelStart = CHARRANGE.cpMin
    rtb.SelLength = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmConfColores = Nothing
    
End Sub



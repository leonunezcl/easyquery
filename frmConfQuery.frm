VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConfQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar editor de Querys"
   ClientHeight    =   3915
   ClientLeft      =   3330
   ClientTop       =   2325
   ClientWidth     =   5955
   Icon            =   "frmConfQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Color Editor"
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
      Index           =   5
      Left            =   4425
      TabIndex        =   8
      Top             =   1860
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir Editor"
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
      Index           =   4
      Left            =   4425
      TabIndex        =   7
      Top             =   2295
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Fondo Editor"
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
      Index           =   3
      Left            =   4410
      TabIndex        =   6
      Top             =   1425
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Fuente Editor"
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
      Index           =   2
      Left            =   4410
      TabIndex        =   5
      Top             =   990
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aplicar Cambios"
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
      Left            =   4410
      TabIndex        =   4
      Top             =   555
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
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
      Left            =   4410
      TabIndex        =   3
      Top             =   120
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview de editor de Querys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   390
      TabIndex        =   1
      Top             =   30
      Width           =   3930
      Begin RichTextLib.RichTextBox rtbQuery 
         Height          =   3450
         Left            =   75
         TabIndex        =   2
         Top             =   270
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   6085
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmConfQuery.frx":08CA
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3825
      Left            =   0
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmConfQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Cdlg As New GCommonDialog
Private mGradient As New clsGradient
Private Sub Aplicar()

    Dim f As New frmQuery
    Dim ArrayCount As Integer
    Dim k As Integer
    
    Call GrabaIni(C_INI, "Editor", "Letra", rtbQuery.Font.Name)
    Call GrabaIni(C_INI, "Editor", "Size", rtbQuery.Font.Size)
    Call GrabaIni(C_INI, "Editor", "BackColor", rtbQuery.BackColor)
    Call GrabaIni(C_INI, "Editor", "ColorFuente", rtbQuery.SelColor)
    
    glbFontLetra = rtbQuery.Font.Name
    glbFontSize = rtbQuery.Font.Size
    glbBackColor = rtbQuery.BackColor
    glbColorSql = rtbQuery.SelColor
    
    ArrayCount = UBound(Document)

    For k = 1 To ArrayCount
        If Not fState(k).Deleted Then
            Set f = Document(k)
            f.HayCambios = True
            f.txtQuery.SelStart = 0
            f.txtQuery.SelLength = Len(f.txtQuery.Text)
            f.txtQuery.SelFontName = glbFontLetra
            f.txtQuery.SelFontSize = glbFontSize
            f.txtQuery.BackColor = glbBackColor
            f.txtQuery.SelColor = glbColorSql
            f.txtQuery.SelLength = 0
            f.FormateaSentencias
            
        End If
    Next

    Set f = Nothing
            
End Sub

'carga las opciones del editor
Private Sub CargaOpcionesEditor()

    rtbQuery.Text = "SELECT * FROM AUTHORS WHERE AUTHOR_ID = 1000"
    rtbQuery.Font.Name = glbFontLetra
    rtbQuery.Font.Size = glbFontSize
    rtbQuery.BackColor = glbBackColor
    rtbQuery.SelStart = 0
    rtbQuery.SelLength = Len(rtbQuery.Text)
    rtbQuery.SelColor = glbColorSql
    rtbQuery.SelLength = 0
            
End Sub

Private Sub Color()

    Dim lngColor As Long
    
    lngColor = rtbQuery.SelColor
    
    If Cdlg.VBChooseColor(lngColor) Then
        rtbQuery.SelStart = 0
        rtbQuery.SelLength = Len(rtbQuery.Text)
        rtbQuery.SelColor = lngColor
        rtbQuery.SelLength = 0
    End If
    
End Sub

Private Sub Fondo()

    Dim lngColor As Long
    
    lngColor = rtbQuery.BackColor
    
    If Cdlg.VBChooseColor(lngColor) Then
       rtbQuery.BackColor = lngColor
    End If
    
End Sub

Private Sub Fuente()
    
    Dim Fuente As New StdFont
    Dim lngColor As Long
    
    Fuente.Name = rtbQuery.Font.Name
    lngColor = rtbQuery.SelColor
    If Cdlg.VBChooseFont(Fuente) Then
        rtbQuery.Font.Name = Fuente.Name
        rtbQuery.Font.Size = Fuente.Size
        rtbQuery.SelStart = 0
        rtbQuery.SelLength = Len(rtbQuery.Text)
        rtbQuery.SelColor = lngColor
        rtbQuery.SelLength = 0
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'Aceptar
            Call Aplicar
            Unload Me
        Case 1  'Aplicar
            Call Aplicar
        Case 2  'Fuente
            Call Fuente
        Case 3  'Fondo
            Call Fondo
        Case 4  'Salir
            Unload Me
        Case 5
            Call Color
    End Select
        
End Sub

Private Sub Form_Load()

    Hourglass hwnd, True
    
    CenterWindow hwnd
    
    CargaOpcionesEditor
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
        
    picDraw.Refresh
    
    Hourglass hwnd, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConfQuery = Nothing
End Sub



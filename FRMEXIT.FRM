VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información "
   ClientHeight    =   5355
   ClientLeft      =   1350
   ClientTop       =   2235
   ClientWidth     =   5550
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4425
      TabIndex        =   1
      Top             =   4845
      Width           =   915
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5355
      Left            =   0
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   525
      TabIndex        =   7
      Top             =   1155
      Width           =   4950
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Explorer analiza proyectos creados con visual basic 3 , 4, 5 y 6."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   525
      TabIndex        =   6
      Top             =   525
      Width           =   4650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblGlosa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Necesita un analizador de Proyectos ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   540
      TabIndex        =   5
      Top             =   135
      Width           =   3705
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2002 Luis Núñez Ibarra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   495
      MouseIcon       =   "frmExit.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "http://www.vbsoftware/autor.html"
      Top             =   4860
      Width           =   3105
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   495
      MouseIcon       =   "frmExit.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "http://www.vbsoftware.cl"
      Top             =   5085
      Width           =   2370
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Explorer Home Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   495
      MouseIcon       =   "frmExit.frx":0EDE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "http://www.vbsoftware.cl/vbpexplorer.html"
      Top             =   4635
      Width           =   2070
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    Dim msg As String
        
    msg = "Información sobre : " & vbNewLine
    msg = msg & "- Referencias y componentes." & vbNewLine
    msg = msg & "- Formularios, módulos bas, módulos cls , controles de usuario y "
    msg = msg & "página de propiedades." & vbNewLine & vbNewLine
    msg = msg & "Despliega graficamente : " & vbNewLine
    msg = msg & " - subs, funciones, variables, apis , tipos, enumeraciones , propiedades, "
    msg = msg & "eventos , controles y arreglos." & vbNewLine
    msg = msg & "Analiza el proyecto y le informa : " & vbNewLine & vbNewLine
    msg = msg & " - variables no usadas." & vbNewLine
    msg = msg & " - constantes no usadas." & vbNewLine
    msg = msg & " - apis no usadas." & vbNewLine
    msg = msg & " - funciones no usadas." & vbNewLine
    msg = msg & " - subs no usadas." & vbNewLine
    msg = msg & " - filtrar el analisis y archivos a explorar." & vbNewLine
    msg = msg & " - Imprimir reporte de análisis y más ..." & vbNewLine
    
    lblDescrip.Caption = msg
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, "VBSoftware")
    
    picDraw.Refresh
    
    frmMain.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmExit = Nothing
    
End Sub


Private Sub lblCopyright_Click()
    pShell hwnd, lblCopyright.Tag
End Sub


Private Sub lblProduct_Click()
    pShell hwnd, lblProduct.Tag
End Sub


Private Sub lblURL_Click()
    pShell hwnd, lblURL.Tag
End Sub



VERSION 5.00
Begin VB.Form frmAcercaDe 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de ..."
   ClientHeight    =   5280
   ClientLeft      =   3780
   ClientTop       =   2265
   ClientWidth     =   5565
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AcercaDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   5280
      Left            =   0
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4425
      TabIndex        =   0
      Top             =   4860
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   435
      Picture         =   "AcercaDe.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Explorer"
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
      Left            =   3705
      MouseIcon       =   "AcercaDe.frx":1194
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Tag             =   "http://www.vbsoftware.cl/pexplorer.html"
      Top             =   4290
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Query fue explorado y analizado con :"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   495
      TabIndex        =   8
      Top             =   4290
      Width           =   3150
   End
   Begin VB.Label lblGlosa 
      BackStyle       =   0  'Transparent
      Caption         =   "Realiza múltiples consultas a orígenes de datos ODBC y bases de datos Access 97/2000"
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
      Height          =   750
      Left            =   1020
      TabIndex        =   7
      Top             =   210
      Width           =   4095
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
      Left            =   510
      MouseIcon       =   "AcercaDe.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "http://www.vbsoftware.cl/autor.html"
      Top             =   4830
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
      Left            =   510
      MouseIcon       =   "AcercaDe.frx":17A8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://www.vbsoftware.cl"
      Top             =   5055
      Width           =   2370
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Query ! Home Page"
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
      Left            =   510
      MouseIcon       =   "AcercaDe.frx":1AB2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "http://www.vbsoftware.cl/easyquery.html"
      Top             =   4605
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1710
      TabIndex        =   2
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      TabIndex        =   1
      Top             =   990
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    
    Dim Msg As String
    
    CenterWindow hwnd
        
    Msg = "Creado por Luis Núñez Ibarra." & vbNewLine
    Msg = Msg & "Todos los derechos reservados." & vbNewLine
    Msg = Msg & "Santiago de Chile 2001" & vbNewLine & vbNewLine
    Msg = Msg & "Esta aplicación se distribuye libre de cargo alguno "
    Msg = Msg & "y ha sido probada con origenes de datos DB/2 , Sql Server , "
    Msg = Msg & "Access 97 y 2000." & vbNewLine & vbNewLine
    Msg = Msg & "VBSoftware no se hace responsable por algún daño ocasionado "
    Msg = Msg & "por el uso de esta aplicación." & vbNewLine & vbNewLine
    Msg = Msg & "Esta es una versión beta. Tus sugerencias, ideas , reporte "
    Msg = Msg & "de errores y criticas haran de este software "
    Msg = Msg & "una gran ayuda a los desarrolladores."
    
    Label3.Caption = Msg
                
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, App.Title & " Beta Versión : " & App.Major & "." & App.Minor & "." & App.Revision)
    
    picDraw.Refresh
            
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    If Not gbInicio Then
        gbInicio = True
        frmMain.Show
    End If
    
    Set mGradient = Nothing
    Set frmAcercaDe = Nothing
    
End Sub





Private Sub Label4_Click()
    pShell Label4.Tag, hwnd
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



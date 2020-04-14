VERSION 5.00
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   4935
   ClientLeft      =   2445
   ClientTop       =   2565
   ClientWidth     =   7485
   Icon            =   "Imprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   25
      Top             =   0
      Width           =   360
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
      Index           =   1
      Left            =   4035
      TabIndex        =   24
      Top             =   4470
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
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
      Left            =   2355
      TabIndex        =   23
      Top             =   4470
      Width           =   1245
   End
   Begin VB.Frame fra 
      Caption         =   "Parámetros de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   435
      TabIndex        =   0
      Top             =   90
      Width           =   6975
      Begin VB.CheckBox chkLineas 
         Caption         =   "Imprimir lineas consulta"
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   3870
         Width           =   2565
      End
      Begin VB.Frame fraMargenes 
         Caption         =   "Márgenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   150
         TabIndex        =   13
         Top             =   2640
         Width           =   6675
         Begin VB.TextBox txtDer 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            TabIndex        =   21
            Text            =   "1"
            Top             =   750
            Width           =   1005
         End
         Begin VB.TextBox txtIzq 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            TabIndex        =   19
            Text            =   "1"
            Top             =   330
            Width           =   1005
         End
         Begin VB.TextBox txtAbajo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            TabIndex        =   17
            Text            =   "1"
            Top             =   750
            Width           =   1005
         End
         Begin VB.TextBox txtArriba 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            TabIndex        =   15
            Text            =   "1"
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Derecha"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3150
            TabIndex        =   20
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierda"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3150
            TabIndex        =   18
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Abajo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   16
            Top             =   780
            Width           =   405
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Arriba"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   14
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.ComboBox cboOrientacion 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Frame fraPaginas 
         Caption         =   "Rango de Páginas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   105
         TabIndex        =   6
         Top             =   1395
         Width           =   6735
         Begin VB.TextBox txtHasta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2940
            TabIndex        =   10
            Text            =   "1"
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtDesde 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   930
            TabIndex        =   8
            Text            =   "1"
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2250
            TabIndex        =   9
            Top             =   450
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   420
            Width           =   465
         End
      End
      Begin VB.Frame fraEncabezados 
         Caption         =   "Encabezados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6765
         Begin VB.TextBox txtPie 
            Height          =   315
            Left            =   1170
            TabIndex        =   5
            Top             =   600
            Width           =   5505
         End
         Begin VB.TextBox txtCabezera 
            Height          =   315
            Left            =   1170
            TabIndex        =   4
            Top             =   240
            Width           =   5505
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pie de Página"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cabezera"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Orientacion"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2310
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Sub Imprimir()

    On Local Error GoTo ErrorImprimir
    
    Call Hourglass(hwnd, True)
    
    If Not frmPreview.Estructura Then
        frmMain.ActiveForm.griQuery.PrintHeader = txtCabezera.Text
        frmMain.ActiveForm.griQuery.PrintFooter = txtPie.Text
        frmMain.ActiveForm.griQuery.PrintPageStart = Val(txtDesde.Text)
        frmMain.ActiveForm.griQuery.PrintPageEnd = Val(txtHasta.Text)
        frmMain.ActiveForm.griQuery.PrintOrientation = cboOrientacion.ListIndex
        frmMain.ActiveForm.griQuery.PrintMarginTop = Val(txtArriba.Text)
        frmMain.ActiveForm.griQuery.PrintMarginBottom = Val(txtAbajo.Text)
        frmMain.ActiveForm.griQuery.PrintMarginLeft = Val(txtIzq.Text)
        frmMain.ActiveForm.griQuery.PrintMarginRight = Val(txtDer.Text)
        frmMain.ActiveForm.griQuery.PrintGrid = chkLineas.Value
        frmMain.ActiveForm.griQuery.PrintSmartPrint = True
        frmMain.ActiveForm.griQuery.Action = 32
    Else
        frmVerTabla.fpTabla.PrintHeader = txtCabezera.Text
        frmVerTabla.fpTabla.PrintFooter = txtPie.Text
        frmVerTabla.fpTabla.PrintPageStart = Val(txtDesde.Text)
        frmVerTabla.fpTabla.PrintPageEnd = Val(txtHasta.Text)
        frmVerTabla.fpTabla.PrintOrientation = cboOrientacion.ListIndex
        frmVerTabla.fpTabla.PrintMarginTop = Val(txtArriba.Text)
        frmVerTabla.fpTabla.PrintMarginBottom = Val(txtAbajo.Text)
        frmVerTabla.fpTabla.PrintMarginLeft = Val(txtIzq.Text)
        frmVerTabla.fpTabla.PrintMarginRight = Val(txtDer.Text)
        frmVerTabla.fpTabla.PrintGrid = chkLineas.Value
        frmVerTabla.fpTabla.PrintSmartPrint = True
        frmVerTabla.fpTabla.Action = 32
    End If
    
    MsgBox "Impresión realizada.", vbInformation
    
    GoTo SalirImprimir
    
ErrorImprimir:
    MsgBox "Imprimir : " & Err & " " & Error$, vbCritical
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    Select Case Index
        Case 0
            Call Imprimir
        Case 1
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()

    cboOrientacion.AddItem "DEFAULT"
    cboOrientacion.AddItem "PORTRAIT"
    cboOrientacion.AddItem "LANDSCAPE"
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, "Imprimir")
    
    picDraw.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmImprimir = Nothing
    Set mGradient = Nothing
    
End Sub



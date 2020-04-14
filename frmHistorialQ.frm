VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistorialQ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Querys"
   ClientHeight    =   5700
   ClientLeft      =   3210
   ClientTop       =   2340
   ClientWidth     =   7245
   Icon            =   "frmHistorialQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Editor de historial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   405
      TabIndex        =   7
      Top             =   1275
      Width           =   5385
      Begin RichTextLib.RichTextBox rtbQuery 
         Height          =   2085
         Left            =   135
         TabIndex        =   10
         Top             =   2175
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   3678
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmHistorialQ.frx":08CA
      End
      Begin MSComctlLib.ListView lvwHis 
         Height          =   1680
         Left            =   105
         TabIndex        =   8
         Top             =   255
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2963
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sql"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sentencia SQL digitada"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1965
         Width           =   2025
      End
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
      Left            =   5880
      TabIndex        =   6
      Top             =   570
      Width           =   1215
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
      Left            =   5865
      TabIndex        =   5
      Top             =   135
      Width           =   1215
   End
   Begin VB.Frame fraH 
      Caption         =   "Opciones de historial de Querys realizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   390
      TabIndex        =   1
      Top             =   15
      Width           =   5400
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2715
         MaxLength       =   3
         TabIndex        =   12
         Top             =   840
         Width           =   750
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   4
         Top             =   465
         Width           =   330
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   4845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "sql realizados"
         Height          =   195
         Index           =   1
         Left            =   3510
         TabIndex        =   13
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de historial contiene hasta:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   855
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path del archivo de historial"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   255
         Width           =   1950
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5625
      Left            =   0
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistorialQ.frx":0938
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistorialQ.frx":0A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistorialQ.frx":0BF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistorialQ.frx":0D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistorialQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Sub CambiarPath()

    glbMisDocumentos = txtPath.Text
    
    CopyFile txtPath.Tag & "easyquery.his", txtPath.Text & "easyquery.his", 0
    DeleteFile txtPath.Tag & "easyquery.his"
    Call GrabaIni(C_INI, "Historial", "Path", glbMisDocumentos)
        
    MsgBox "Configuración actualizada con éxito!", vbInformation
    
End Sub


Private Function MyIsnumeric(ByVal Numero As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    ret = True
    
    For k = 1 To Len(Numero)
        Select Case Mid$(Numero, k, 1)
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            Case Else
                ret = False
                Exit For
        End Select
    Next k
    
    MyIsnumeric = ret
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If txtNumero.Text <> txtNumero.Tag Then
            If MyIsnumeric(txtNumero.Text) Then
                Call GrabaIni(C_INI, "Historial", "Numero", txtNumero.Text)
                glbNumSQl = txtNumero.Text
            End If
        End If
        
        If txtPath.Text <> txtPath.Tag Then
            Call CambiarPath
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub cmdPath_Click()

    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = Me.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    If sPath <> "" Then
        If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
        txtPath.Text = sPath
    End If
    
End Sub


Private Sub Form_Load()

    Hourglass hwnd, True
    
    CenterWindow hwnd
                
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    txtPath.Text = glbMisDocumentos
    txtPath.Tag = txtPath.Text
    txtNumero.Text = glbNumSQl
    txtNumero.Tag = glbNumSQl
    
    Call CargaHistorial(Me.lvwHis)
    
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
    
    Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmHistorialQ = Nothing
End Sub


Private Sub lvwHis_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rtbQuery.Text = Item.SubItems(1)
End Sub



VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form frmConfSkin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Skin"
   ClientHeight    =   3585
   ClientLeft      =   2835
   ClientTop       =   2535
   ClientWidth     =   4260
   Icon            =   "ConfSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.SkinForm SkinForm 
      Height          =   480
      Left            =   2970
      OleObjectBlob   =   "ConfSkin.frx":08CA
      TabIndex        =   5
      Top             =   2340
      Width           =   480
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Salir"
      Height          =   465
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   1380
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aplicar"
      Height          =   465
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
      Height          =   465
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Top             =   300
      Width           =   1245
   End
   Begin VB.ListBox lisSkin 
      Height          =   3180
      ItemData        =   "ConfSkin.frx":09A1
      Left            =   60
      List            =   "ConfSkin.frx":09ED
      TabIndex        =   1
      Top             =   270
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Skin Disponibles"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   1170
   End
End
Attribute VB_Name = "frmConfSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WinAmpXML As String
Private Sub AplicaSkin()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim F As Form
    
    If UBound(Document) = 0 Then Exit Sub
    
    If lisSkin.Text <> "ninguno" Then
        glbPathSkin = App.Path & "\skins\" & lisSkin.Text & ".skn"
        Call Uninstall(frmMain.hWnd)
    Else
        glbPathSkin = ""
        Call Uninstall(frmMain.hWnd)
        Call frmMain.InstalaSistemaAyuda
    End If
    
    Dim I As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)
    
    frmMain.SkinForm.SkinPath = glbPathSkin
    
    For I = 1 To ArrayCount
        If Not fState(I).Deleted Then
            Set F = Document(I)
        
            F.SkinForm.SkinPath = glbPathSkin
        
            DoEvents
        End If
    Next
    
    Set F = Nothing
    
    Err = 0
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    Select Case Index
        Case 0  'Aceptar
            Call AplicaSkin
            Unload Me
        Case 1  'Aplicar
            Call AplicaSkin
        Case 2  'Salir
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    
    On Local Error Resume Next
    
    CenterWindow hWnd
    
    SkinForm.SkinPath = glbPathSkin 'App.Path + "\Skins"
    
    Err = 0
    
End Sub


Private Sub lisSkin_Click()
    
    If lisSkin.Text <> "ninguno" Then
        SkinForm.SkinPath = App.Path & "\skins\" & lisSkin.Text & ".skn"
    Else
        SkinForm.SkinPath = ""
    End If
    
End Sub



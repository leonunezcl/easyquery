VERSION 5.00
Begin VB.Form frmPaso 
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   2745
   ClientTop       =   1545
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   1740
   Begin VB.PictureBox picPegar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      Picture         =   "paso.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picCortar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   720
      Picture         =   "paso.frx":0166
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picCopiar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "paso.frx":02CC
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picConectar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "paso.frx":0432
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picDesconectar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   330
      Picture         =   "paso.frx":0598
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picNuevo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "paso.frx":06FE
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picBorrar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   330
      Picture         =   "paso.frx":0864
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picAbrir 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   30
      Picture         =   "paso.frx":09CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picGrabar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   330
      Picture         =   "paso.frx":0ACC
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   1110
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picImprimir 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   660
      Picture         =   "paso.frx":0C32
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


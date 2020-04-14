VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "VisData Database Utility"
'>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>

Public VBInstance As VBIDE.VBE

Implements IDTExtensibility

'---------------------------------------------------
'this method is used as a call from an external
'launch utility or the VB IDE
'---------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, _
        ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
        
  On Error GoTo LVDErr
    
  Dim rc As Long
  
  'Set gVDClass = Me
  'this sets the VB instance handle
  'that is used by the DataFormDesigner form
  Set VBInstance = VBInst
  frmMain.WindowState = vbNormal
  frmMain.SetFocus
  
  Exit Sub
    
LVDErr:
  MsgBox "IDTExtensibility_OnConnection : " & Err & " " & Err.Description, vbCritical
  
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
  On Error Resume Next
  Unload frmMain
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'---------------------------------------------------
'this method simply closes visdata from the client
'---------------------------------------------------
Sub VDClose()
  Unload frmMain
End Sub


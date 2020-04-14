Attribute VB_Name = "mExportar"
Option Explicit

Private xConnection As ADODB.Connection
Private Rs As New ADODB.Recordset
Public Function CSVExport(ByVal sDest As String) As Boolean
   
    Dim nI            As Long
    Dim nJ            As Long
    Dim nFile         As Integer
    Dim sTmp          As String
    Dim Campo
    Dim Sel
    
    On Error GoTo Err_Handler
      
    ' *** Open output file
    nFile = FreeFile
   
    Open sDest For Output As #nFile
   
    '*** Export fields name
    For nI = 2 To frmMain.ActiveForm.griQuery.MaxCols
       Call frmMain.ActiveForm.griQuery.GetText(nI, 0, Campo)
       sTmp = "" & Campo
       Write #nFile, sTmp;
    Next
    Write #nFile,
    
    For nI = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.GetText(1, nI, Sel)
        
        If Sel = "1" Then
            For nJ = 2 To frmMain.ActiveForm.griQuery.MaxCols
              Call frmMain.ActiveForm.griQuery.GetText(nJ, nI, Campo)
              sTmp = "" & (Campo)
              Write #nFile, sTmp;
            Next
            
            Write #nFile,
        End If
    Next
   
   Close #nFile
   
   CSVExport = True
   
   Exit Function
   
Err_Handler:
   MsgBox ("CSVExport : " & Err & " " & Err.Description)

   CSVExport = False
   
End Function
Public Function EXCELExport(ByVal Archivo As String) As Boolean

    EXCELExport = frmMain.ActiveForm.griQuery.ExportToExcel(Archivo, "hoja 1", App.Path & "\LOGFILE.TXT")
    
End Function

Public Function HTMLExport(ByVal Archivo As String) As Boolean
        
    HTMLExport = frmMain.ActiveForm.griQuery.ExportToHTML(Archivo, False, App.Path & "\LOGFILE.TXT")
    
End Function

'exportar a archivos de tabulaciones
Public Function TABExport(ByVal sDest As String) As Boolean

    Dim nI            As Long
    Dim nJ            As Long
    Dim nFile         As Integer
    Dim sTmp          As String
    Dim Campo
    Dim Sel
    Dim Linea As String
    
    On Error GoTo Err_Handler
      
    ' *** Open output file
    nFile = FreeFile
   
    Open sDest For Output As #nFile
   
   Linea = ""
    '*** Export fields name
    For nI = 2 To frmMain.ActiveForm.griQuery.MaxCols
       Call frmMain.ActiveForm.griQuery.GetText(nI, 0, Campo)
       sTmp = "" & Campo
       
        Linea = Linea & Campo & vbTab
    Next
    
    Print #nFile, Linea
    Linea = ""
    For nI = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.GetText(1, nI, Sel)
        Linea = ""
        If Sel = "1" Then
            For nJ = 2 To frmMain.ActiveForm.griQuery.MaxCols
              Call frmMain.ActiveForm.griQuery.GetText(nJ, nI, Campo)
              sTmp = "" & (Campo)
              
                Linea = Linea & Campo & vbTab
            Next
            
            Print #nFile, Linea
        End If
    Next
   
    Close #nFile
   
    TABExport = True
   
    Exit Function
   
Err_Handler:
   MsgBox "TABExport : " & Err & " " & Err.Description, vbCritical

   TABExport = False
    
End Function
Public Function TXTExport(ByVal sDest As String) As Boolean

    Dim nI            As Long
    Dim nJ            As Long
    Dim nFile         As Integer
    Dim sTmp          As String
    Dim Campo
    Dim Sel
    
    On Error GoTo Err_Handler
      
    ' *** Open output file
    nFile = FreeFile
   
    Open sDest For Output As #nFile
   
    '*** Export fields name
    For nI = 2 To frmMain.ActiveForm.griQuery.MaxCols
       Call frmMain.ActiveForm.griQuery.GetText(nI, 0, Campo)
       sTmp = "" & Campo
       
        If nI < frmMain.ActiveForm.griQuery.MaxCols Then
            Print #nFile, sTmp; ",";
        Else
            Print #nFile, sTmp;
        End If
    Next
    
    Print #nFile,
    
    For nI = 1 To frmMain.ActiveForm.griQuery.DataRowCnt
        Call frmMain.ActiveForm.griQuery.GetText(1, nI, Sel)
        
        If Sel = "1" Then
            For nJ = 2 To frmMain.ActiveForm.griQuery.MaxCols
              Call frmMain.ActiveForm.griQuery.GetText(nJ, nI, Campo)
              sTmp = "" & (Campo)
              
                If nJ < frmMain.ActiveForm.griQuery.MaxCols Then
                    Print #nFile, sTmp; ",";
                Else
                    Print #nFile, sTmp;
                End If
            Next
            
            Print #nFile,
        End If
    Next
   
    Close #nFile
   
    TXTExport = True
   
    Exit Function
   
Err_Handler:
   MsgBox "TXTExport : " & Err & " " & Err.Description, vbCritical

   TXTExport = False
   
End Function

Public Function XMLExport(ByVal Archivo As String) As Boolean

    On Local Error GoTo ErrorXMLExport
    
    Dim ret As Boolean
    Dim sql As String
    'Dim Nombre
    'Dim sCampo As String
    Dim ActiveConexion As Integer
    Dim ci As ComboItem
    
    If Not Conectado Then
        MsgBox "Debe conectarse primero.", vbCritical
        Exit Function
    End If
    
    sql = Trim$(frmMain.ActiveForm.txtQuery.Text)
    
    If sql = "" Then Exit Function
    
    If frmMain.ActiveForm.m_Ejecutando Then Exit Function
    
    If frmMain.ActiveForm.imgConexiones.SelectedItem Is Nothing Then
        MsgBox "Seleccione una conexión.", vbCritical
        frmMain.ActiveForm.imgConexiones.SetFocus
        Exit Function
    End If
    
    Call Hourglass(frmMain.hwnd, True)
    
    ret = True
    
    Set ci = frmMain.ActiveForm.imgConexiones.SelectedItem
    
    ActiveConexion = ConexionActiva(ci.Text)
    
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Grabando archivo ...."
        
    Set xConnection = DBConnection(ActiveConexion)
    xConnection.CommandTimeout = glbTimeOut
    
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
        
    Rs.Open sql, xConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    Rs.Save Archivo, adPersistXML
    
    Rs.Close
    
    GoTo SalirXMLExport
    
ErrorXMLExport:
    ret = False
    MsgBox "XMLExport : " & Err & " " & Error$, vbCritical
    Resume SalirXMLExport
    
SalirXMLExport:
    Call Hourglass(frmMain.hwnd, False)
    frmMain.ActiveForm.staQuery.Panels(2).Text = "Listo."
    XMLExport = ret
    Err = 0
    
End Function

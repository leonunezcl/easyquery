VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Detener As Boolean
Dim Db As New Connection
Dim Rs As New Recordset
Public Sub EjecutaQuery(ByVal Conect As Connection, ByVal hWnd As Long, ByVal NumQuery As Integer)

    On Local Error GoTo ErrorEjecutaQuery
    
    Dim Msg As String
    Dim Sql As String
    Dim Campo As Field
    Dim Ncampos As Integer
    Dim Nombre
    Dim k As Long
    Dim f As Integer
    Dim e As Integer
    
    If Not Conectado Then Exit Sub
        
    Sql = Trim$(frmQuery.txtQuery(NumQuery).Text)
    
    If Sql = "" Then Exit Sub
    
    frmQuery.griQuery(NumQuery).MaxRows = 1
    frmQuery.griQuery(NumQuery).MaxCols = 1
    Call frmQuery.griQuery(NumQuery).SetText(1, 0, "")
    Call frmQuery.griQuery(NumQuery).SetText(1, 1, "")
    
    Call Hourglass(hWnd, True)
    
    If Left$(UCase$(Sql), 6) = "SELECT" Then
            
        Rs.CursorLocation = adUseClient
        Rs.CursorType = adOpenStatic
        Set Rs = Conect.Execute(Sql)
        
        Ncampos = 1
        If Not Rs.EOF Then
        With Rs
            For Each Campo In .Fields
                If Ncampos > frmQuery.griQuery(NumQuery).MaxCols Then
                    frmQuery.griQuery(NumQuery).MaxCols = frmQuery.griQuery(NumQuery).MaxCols + 1
                End If
                Nombre = Campo.Name
                Call frmQuery.griQuery(NumQuery).SetText(Ncampos, 0, Nombre)
                Ncampos = Ncampos + 1
            Next
        
            f = 1
            Do While Not .EOF
                e = DoEvents()
                If m_Detener Then Exit Do
                If f Mod 100 = 0 Then Debug.Print NumQuery
                If frmQuery.griQuery(NumQuery).DataRowCnt >= frmQuery.griQuery(NumQuery).MaxRows Then
                    frmQuery.griQuery(NumQuery).MaxRows = frmQuery.griQuery(NumQuery).MaxRows + 1
                End If
                For k = 1 To Ncampos - 1
                    Call frmQuery.griQuery(NumQuery).SetText(k, f, .Fields(k - 1))
                Next k
                f = f + 1
                .MoveNext
            Loop
        End With
        End If
        
        Rs.Close
        
        Set Rs = Nothing
    Else
        Call Conect.Execute(Sql)
    End If
                    
    Call Hourglass(hWnd, False)
    
    frmQuery.staBar.Panels(1).Text = "Ok"
    
    Exit Sub
    
ErrorEjecutaQuery:
    frmQuery.staBar.Panels(NumQuery).Text = Error$
    MsgBox "EjecutaQuery : " & Err & " " & Error$, vbCritical
    Err = 0
    Resume Salir
    
Salir:
    Call Hourglass(hWnd, False)
    
End Sub


Public Property Get Detener() As Boolean
    Detener = m_Detener
End Property

Public Property Let Detener(ByVal pDetener As Boolean)
    m_Detener = pDetener
End Property


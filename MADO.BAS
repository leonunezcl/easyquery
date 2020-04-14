Attribute VB_Name = "mAdo"
Option Explicit

'****************************
'creado por Luis Núñez Ibarra
'
'Módulo con funciones utiles de ADO
'
'URL
'http://proyectovb.topcities.com
'
'EMAIL
'proyectovb@hotmail.com
'****************************

Public DBConnection() As New ADODB.Connection
Public DBRecordset As New ADODB.Recordset
Public DBCommand As New ADODB.Command
'Abre una conexion segun un origen de datos ODBC
'definido en el panel de control ODBC
'
'True = Exito al conectar
'False = Error al conectar
Public Function AbrirConexion(ByVal Conexion As String, _
                               Optional ByVal Usuario As String = "", _
                               Optional Password As String = "") As Boolean

    On Error GoTo ErrorAbrirConexion
    
    Dim ret As Boolean
    
    ret = True
    
    DBConnection.ConnectionTimeout = 0 'no hay limite de espera de tiempo
    
    If Usuario <> "" Then
        DBConnection.Open Conexion, Usuario, Password
    Else
        DBConnection.Open Conexion
    End If
    
    GoTo SalirAbrirConexion
     
ErrorAbrirConexion:
    ret = False
    MsgBox "AbrirConexion : " & Err & " " & Error$, vbCritical
    Resume SalirAbrirConexion
    
SalirAbrirConexion:
    AbrirConexion = ret
    Err = 0
    
End Function


'Abre una base de datos access
'Regresa :
'True = Exito al conectar
'False = Error al conectar
'
Public Function AbrirBaseDatosAccess(ByVal BaseAccess As String, _
                               Optional ByVal Access2000 As Boolean = False, _
                               Optional ByVal Usuario As String = "", _
                               Optional ByVal Password As String = "") As Boolean

    On Error GoTo ErrorAbrirBaseDatosAccess
    
    Dim ret As Boolean
    Dim Conexion As String
    
    ret = True
    
    If Not Access2000 Then 'access 97
        Conexion = "Provider=Microsoft.Jet.OLEDB.3.51;" _
            & "Persist Security Info=False;Data Source=" & BaseAccess
    Else                    'access 2000
        Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Persist Security Info=False;Data Source=" & BaseAccess
    End If
        
    DBConnection.ConnectionTimeout = 0 'no hay limite de espera de tiempo
    
    If Usuario <> "" Then
        DBConnection.Open Conexion, Usuario, Password
    Else
        DBConnection.Open Conexion
    End If
    
    GoTo SalirAbrirBaseDatosAccess
     
ErrorAbrirBaseDatosAccess:
    ret = False
    MsgBox "AbrirBaseDatosAccess : " & Err & " " & Error$, vbCritical
    Resume SalirAbrirBaseDatosAccess
    
SalirAbrirBaseDatosAccess:
    AbrirBaseDatosAccess = ret
    Err = 0
                
End Function


'abre un adodb.recordset
'parametros :
'   adors     = un recordset a abrir
'   sqlstmt   = la cade sql con el select realizado
'   modoabrir =
Public Function AbrirRecordset(ByVal AdoRs As ADODB.Recordset, ByVal SqlStmt As String, _
                                ModoAbrir As eOpenRecordset) As Boolean

    On Error GoTo ErrorAbrirRecordset
    
    Dim ret As Boolean
    
    ret = True
    
    If ModoAbrir = OPEN_SERVIDOR_TURBO Then
    
        'Abre el cursor en modo solo avance ->
        'no se puede realizar MoveFirst/MoveNext/MovePrevius/MoveLast/Move/Recordcount
        
        '<es la forma mas rapida de abrir un cursor>.
        'útil para cuando queremos por ejemplo :
        'llenar una grilla, combo, lista, etc
        
        DBRecordset.CursorLocation = adUseServer
        DBRecordset.Open SqlStmt, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    ElseIf ModoAbrir = OPEN_CLIENTE_TURBO Then
        
        'Abre el cursor en modo solo avance ->
        'no se puede realizar MoveFirst/MoveNext/MovePrevius/MoveLast/Move/Recordcount
                
        'este modo es un poco mas lento. pero solo cambia donde abrimos el cursor
        'tambien es util para cuando queremos por ejemplo :
        'llenar una grilla, combo, lista, etc
        
        DBRecordset.CursorLocation = adUseClient
        DBRecordset.Open SqlStmt, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    ElseIf ModoAbrir = OPEN_CLIENTE_BROWSER Then
        
        'Abre el cursor en modo browser
        'podemos realizar MoveFirst/MoveNext/MovePrevius/MoveLast/Move/Recordcount
        
        'este modo es lento. pero solo cambia donde abrimos el cursor
                        
        DBRecordset.CursorLocation = adUseClient
        DBRecordset.Open SqlStmt, DBConnection, adOpenDynamic, adLockReadOnly, adCmdText
    
    Else
        ret = False 'modo incorrecto
    End If
    
    GoTo SalirAbrirRecordset
    
ErrorAbrirRecordset:
    ret = False
    MsgBox "AbrirRecordset : " & Err & " " & Error$, vbCritical
    Resume SalirAbrirRecordset
    
SalirAbrirRecordset:
    AbrirRecordset = ret
    Err = 0
    
End Function


'agrega parametros delete
Private Sub AgregaParametrosDelete()

    Dim SqlStmt As String
    
    SqlStmt = "DELETE "
    SqlStmt = SqlStmt & "FROM AUTHORS WHERE "
    SqlStmt = SqlStmt & "AU_ID = ? "
    
    With DBCommand
        .Parameters.Append .CreateParameter("AU_ID", adInteger, adParamInput, 5, 0)
        .CommandText = SqlStmt
        .ActiveConnection = DBConnection
        .Prepared = True 'prepara nuestro cursor
    End With
    
End Sub

'prepara el comando para ejecutar un insert
'
Public Sub AgregaParametrosInsert()

    Dim SqlStmt As String
    
    SqlStmt = "INSERT INTO "
    SqlStmt = SqlStmt & "AUTHORS VALUES (? , ? , ? )"
        
    With DBCommand
        .Parameters.Append .CreateParameter("AU_ID", adInteger, adParamInput, 4, 0)
        .Parameters.Append .CreateParameter("AUTHOR", adChar, adParamInput, 50, 0)
        .Parameters.Append .CreateParameter("YEAR_BORN", adInteger, adParamInput, 2, 0)
        
        .CommandText = SqlStmt
        .ActiveConnection = DBConnection
        .Prepared = True 'prepara nuestro cursor
    End With
    
End Sub

'agrega los parametros para realizar un select a la tabla authors en biblio.mdb
Private Sub AgregaParametrosSelectAuthors()

    Dim SqlStmt As String
    
    SqlStmt = "SELECT "
    SqlStmt = SqlStmt & "* FROM AUTHORS WHERE "
    SqlStmt = SqlStmt & "AU_ID = ? "
    
    With DBCommand
        .Parameters.Append .CreateParameter("AU_ID", adInteger, adParamInput, 5, 0)
        .CommandText = SqlStmt
        .ActiveConnection = DBConnection
        .Prepared = True 'prepara nuestro cursor
    End With
    
End Sub

'configura los parametros para ejecutar un update en tabla authors
Private Sub AgregaParametrosUpdate()

    Dim SqlStmt As String
    
    SqlStmt = "UPDATE "
    SqlStmt = SqlStmt & " AUTHORS "
    SqlStmt = SqlStmt & " SET "
    SqlStmt = SqlStmt & " AUTHOR = ? , "
    SqlStmt = SqlStmt & " YEAR_BORN = ? "
    SqlStmt = SqlStmt & " WHERE "
    SqlStmt = SqlStmt & " AU_ID = ? "
    
    With DBCommand
        .Parameters.Append .CreateParameter("AUTHOR", adChar, adParamInput, 50, "")
        .Parameters.Append .CreateParameter("YEAR_BORN", adInteger, adParamInput, 4, 0)
        .Parameters.Append .CreateParameter("AU_ID", adInteger, adParamInput, 4, 0)
        
        .CommandText = SqlStmt
        .ActiveConnection = DBConnection
        .Prepared = True 'prepara nuestro cursor
    End With
    
End Sub

'ejecuta el delete en tabla authors
Private Sub EjecutaDeleteAuthors()

    Dim k As Integer
    Dim Arr_Authors() As eAuthors
    
    ReDim Preserve Arr_Authors(3)
    
    Arr_Authors(1).AU_ID = 1
    Arr_Authors(2).AU_ID = 2
    Arr_Authors(3).AU_ID = 3
        
    'RAPIDISIMO >:-> !
    For k = 1 To UBound(Arr_Authors)
        With DBCommand
            .Parameters("AU_ID") = Arr_Authors(k).AU_ID
                        
            Set DBRecordset = Nothing
            Set DBRecordset = .Execute
        End With
    Next k
    
    'LENTO >:-< !!!
    For k = 1 To UBound(Arr_Authors)
        SqlStmt = "DELETE FROM AUTHORS "
        SqlStmt = SqlStmt & " WHERE "
        SqlStmt = SqlStmt & " AU_ID = " & Arr_Authors(k).AU_ID
        
        DBConnection.Execute SqlStmt
    Next k
    
End Sub
'ejecuta insert en authors
'
Private Sub EjecutaInsertAuthors()

    Dim k As Integer
    Dim Arr_Authors() As eAuthors
    
    ReDim Preserve Arr_Authors(3)
    
    Arr_Authors(1).AU_ID = 1
    Arr_Authors(1).AUTHOR = "MIGUEL DE CERVANTES"
    Arr_Authors(1).YEAR_BORN = 1860
    
    Arr_Authors(2).AU_ID = 2
    Arr_Authors(2).AUTHOR = "STEPHEN KING"
    Arr_Authors(2).YEAR_BORN = 1940
    
    Arr_Authors(3).AU_ID = 3
    Arr_Authors(3).AUTHOR = "ISABEL ALLENDE"
    Arr_Authors(3).YEAR_BORN = 1953
    
    'RAPIDISIMO >:-> !
    For k = 1 To UBound(Arr_Authors)
        With DBCommand
            .Parameters("AU_ID") = Arr_Authors(k).AU_ID
            .Parameters("AUTHOR") = Arr_Authors(k).AUTHOR
            .Parameters("YEAR_BORN") = Arr_Authors(k).YEAR_BORN
            
            Set DBRecordset = Nothing
            Set DBRecordset = .Execute
        End With
    Next k
    
    'LENTO >:-< !!!
    For k = 1 To UBound(Arr_Authors)
        SqlStmt = "INSERT INTO AUTHORS VALUES "
        SqlStmt = SqlStmt & Arr_Authors(k).AU_ID & " , "
        SqlStmt = SqlStmt & "'" & Arr_Authors(k).AUTHOR & "' , "
        SqlStmt = SqlStmt & Arr_Authors(k).YEAR_BORN & ")"
        
        DBConnection.Execute SqlStmt
    Next k
    
End Sub
'ejecuta el select de consulta en tabla authors
Private Sub EjecutaSelectAuthors()

    'se configura el comando pero el ejecute debe establecerse
    'a un recordset
    'lo ideal es preparar el o los comandos globales una vez en el programa
    'y limpiarlos al finalizarlo.
    '
    'si el comando va a ser usado en un formulario preparar parametros
    'al inicio y luego limpiarlos en el unload
    
    With DBCommand
        .Parameters("AU_ID") = 1000
        Set DBRecordset = Nothing
        Set DBRecordset = .Execute
    End With
    
    With DBRecordset
        Do While Not .EOF
            '.........
            .MoveNext
        Loop
    End With
    
    DBRecordset.Close
    
End Sub

'ejecuta el update en tabla authors
Private Sub EjecutaUpdateAuthors()

    Dim k As Integer
    Dim Arr_Authors() As eAuthors
    
    ReDim Preserve Arr_Authors(3)
    
    Arr_Authors(1).AU_ID = 1
    Arr_Authors(1).AUTHOR = "MIGUEL DE CERVANTES Y SAAVEDRA"
    Arr_Authors(1).YEAR_BORN = 1860
    
    Arr_Authors(2).AU_ID = 2
    Arr_Authors(2).AUTHOR = "STEPHEN KING PEREZ"
    Arr_Authors(2).YEAR_BORN = 1940
    
    Arr_Authors(3).AU_ID = 3
    Arr_Authors(3).AUTHOR = "ISABEL ALLENDE GOCE"
    Arr_Authors(3).YEAR_BORN = 1953
    
    'RAPIDISIMO >:-> !
    For k = 1 To UBound(Arr_Authors)
        With DBCommand
            .Parameters("AU_ID") = Arr_Authors(k).AU_ID
            .Parameters("AUTHOR") = Arr_Authors(k).AUTHOR
            .Parameters("YEAR_BORN") = Arr_Authors(k).YEAR_BORN
            
            Set DBRecordset = Nothing
            Set DBRecordset = .Execute
        End With
    Next k
    
    'LENTO >:-< !!!
    For k = 1 To UBound(Arr_Authors)
        SqlStmt = "UPDATE AUTHORS SET "
        SqlStmt = SqlStmt & " AUTHOR = '" & Arr_Authors(k).AUTHOR & "' , "
        SqlStmt = SqlStmt & " YEAR_BORN = " & Arr_Authors(k).YEAR_BORN
        SqlStmt = SqlStmt & " WHERE "
        SqlStmt = SqlStmt & " AU_ID = " & Arr_Authors(k).AU_ID
        
        DBConnection.Execute SqlStmt
    Next k
    
End Sub
'limpiar comando de memoria
Private Sub LimpiarComando()

    DBCommand.Prepared = False
    Set DBCommand = Nothing
    
End Sub



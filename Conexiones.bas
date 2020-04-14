Attribute VB_Name = "mConexiones"
Option Explicit

Public DBRecordset As New ADODB.Recordset

Public Enum eTipoConexiones
    TIPO_ODBC = 1
    TIPO_MDB = 2
End Enum

Type eInfo
    Nombre As String
    Descripción As String
End Type

Type ConectionState
    Conexion As String
    Trx As Boolean
    Deleted As Integer
    tipo As eTipoConexiones
    ArchivoMdb As String
    Conectado As Boolean
    LlaveMdb As String
    aTablas() As eInfo
    aVistas() As eInfo
    aProcs() As eInfo
    aIndice() As eInfo
End Type

Public cState()  As ConectionState
Public DBConnection() As New Connection
Public aArchivosMdb() As String

Type eComandos
    Comando As String
    Activo As Boolean
End Type
Public Comandos() As eComandos



Public Sub CargaConexiones()

    Dim i As Integer
    Dim k As Integer
    Dim f As frmQuery
    Dim ci As ComboItem
    
    Dim ArrayCount As Integer
    Dim ArrayDoc As Integer

    ArrayCount = UBound(DBConnection)
    ArrayDoc = UBound(Document)
    
    For k = 1 To ArrayDoc
        If Not fState(k).Deleted Then
            Set f = Document(k)
            
            f.imgConexiones.ComboItems.Clear
            f.lviewCampos.ListItems.Clear
            f.lviewDetalle.ListItems.Clear
            'f.lvwIndices.ListItems.Clear
            f.lvwProcs.ListItems.Clear
            f.lvwTablas.ListItems.Clear
            f.lvwVistas.ListItems.Clear
            
            For i = 1 To ArrayCount
                If Not cState(i).Deleted Then
                    
                    If cState(i).tipo = TIPO_ODBC Then
                        Set ci = f.imgConexiones.ComboItems.Add(1, cState(i).Conexion, _
                        cState(i).Conexion, 1, 1, 0)
                    Else
                        Set ci = f.imgConexiones.ComboItems.Add(1, cState(i).ArchivoMdb, _
                        cState(i).LlaveMdb, 1, 1, 0)
                    End If
                    
                    ci.Indentation = 2
                    
                    If f.imgConexiones.ComboItems.Count = 1 Then
                        f.imgConexiones.ComboItems(1).Selected = True
                        f.CargaInfoProveedor
                    End If
                End If
            Next i
        End If
    Next k
    
    Set ci = Nothing
    Set f = Nothing
    
End Sub

Function FindFreeConection() As Integer
    
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(DBConnection)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = 1 To ArrayCount
        If cState(i).Deleted Then
            FindFreeConection = i
            cState(i).Deleted = False
            Exit Function
        End If
    Next

    ReDim Preserve DBConnection(ArrayCount + 1)
    ReDim Preserve cState(ArrayCount + 1)
    
    FindFreeConection = UBound(DBConnection)
    
End Function

Public Function NewConection(ByVal Conexion As String, _
                             ByVal TipoConexion As eTipoConexiones, _
                             Optional ByVal Usuario As String = "", _
                             Optional ByVal Password As String = "", _
                             Optional ByVal ArchivoMdb As String = "", _
                             Optional ByVal LlaveMdb As String = "") As Boolean
    
    On Local Error GoTo ErrorNewConection
    
    Dim cIndex As Integer
    Dim ret As Boolean
    Dim Rs As ADODB.Recordset
    Dim k As Integer
    
    ret = True
    
    Load frmInfo
    frmInfo.Show
    DoEvents
    
    cIndex = FindFreeConection()
    
    cState(cIndex).Conexion = Conexion
    cState(cIndex).Deleted = False
    cState(cIndex).tipo = TipoConexion
    cState(cIndex).ArchivoMdb = ArchivoMdb
    cState(cIndex).Conectado = True
    cState(cIndex).LlaveMdb = LlaveMdb
    ReDim cState(cIndex).aTablas(0)
    ReDim cState(cIndex).aVistas(0)
    ReDim cState(cIndex).aProcs(0)
    ReDim cState(cIndex).aIndice(0)
    
    DBConnection(cIndex).ConnectionTimeout = 0
    DBConnection(cIndex).ConnectionString = Conexion
    
    If Usuario <> "" Then
        DBConnection(cIndex).Open , Usuario, Password
    Else
        DBConnection(cIndex).Open Conexion
    End If
        
    'obtener info de la conexion
    'cargar tablas
    Set Rs = New ADODB.Recordset
    Set Rs = DBConnection(cIndex).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    
    k = 1
    Do While Not Rs.EOF
        If Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
            ReDim Preserve cState(cIndex).aTablas(k)
            cState(cIndex).aTablas(k).Nombre = Trim$(Rs!table_name)
            If IsNull(Rs!Description) Then
                cState(cIndex).aTablas(k).Descripción = ""
            Else
                cState(cIndex).aTablas(k).Descripción = Trim$(Rs!Description)
            End If
            k = k + 1
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
        
    'cargar vistas
    Set Rs = DBConnection(cIndex).OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "VIEW"))
    k = 1
    Do While Not Rs.EOF
        If Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
            ReDim Preserve cState(cIndex).aVistas(k)
            cState(cIndex).aVistas(k).Nombre = Trim$(Rs!table_name)
            If IsNull(Rs!Description) Then
                cState(cIndex).aVistas(k).Descripción = ""
            Else
                cState(cIndex).aVistas(k).Descripción = Trim$(Rs!Description)
            End If
            k = k + 1
        End If
        Rs.MoveNext
    Loop
    
    Rs.Close
        
    'cargar procs
    Set Rs = DBConnection(cIndex).OpenSchema(adSchemaProcedures)
    k = 1
    Do While Not Rs.EOF
        ReDim Preserve cState(cIndex).aProcs(k)
        cState(cIndex).aProcs(k).Nombre = Trim$(Rs!procedure_name)
        
        If IsNull(Rs!Description) Then
            cState(cIndex).aProcs(k).Descripción = ""
        Else
            cState(cIndex).aProcs(k).Descripción = Trim$(Rs!Description)
        End If
            
        Rs.MoveNext
        k = k + 1
    Loop
    
    Rs.Close
        
    'cargar indices
    'Set Rs = DBConnection(cIndex).OpenSchema(adSchemaIndexes)
    'k = 1
    'Do While Not Rs.EOF
    '   If Left$(UCase$(Rs!table_name), 4) <> "MSYS" Then
    '        ReDim Preserve cState(cIndex).aIndice(k)
    '        cState(cIndex).aIndice(k).Nombre = Trim$(Rs!INDEX_NAME)
            
            'If IsNull(Rs!PRIMARY_KEY) Then
            '    cState(cIndex).aIndice(K).Descripción = ""
            'Else
            '    cState(cIndex).aIndice(K).Descripción = Trim$(Rs!Description)
            'End If
                
    '        k = k + 1
    '    End If
    '    Rs.MoveNext
    'Loop
    
    'Rs.Close
        
    GoTo SalirNewConection
    
ErrorNewConection:
    Unload frmInfo
    ret = False
    cState(cIndex).Deleted = True
    MsgBox "NewConection : " & Err & " " & Error$, vbCritical
    Resume SalirNewConection
    
SalirNewConection:
    Unload frmInfo
    Set Rs = Nothing
    NewConection = ret
    Err = 0
    
End Function


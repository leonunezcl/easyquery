Attribute VB_Name = "mMdi"
Option Explicit

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
    Tag As Variant
End Type

Type TablasState
    Deleted As Integer
    Dirty As Integer
    Color As Long
    Tag As Variant
End Type


Public fState()  As FormState
Public Document() As New frmQuery
Public fTablasState As TablasState
Public Tablas() As New frmTablas

Public gbMatchCase As Integer
Public gbWholeWord As Integer
Public gsFindText As String
Public gbLastPos As Integer
Public glbFindSql As String
Sub FindText()

    On Local Error GoTo SalirFindText
    
    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim Sql As String
    'Dim iComp As Integer
    
    'If gbMatchCase = 0 Then iComp = 1 Else iComp = 0
        
    Sql = UCase$(glbFindSql)
    
    If gbLastPos = 0 Or gbLastPos > Len(Sql) Then
        lPos = 1
    Else
        lPos = gbLastPos
    End If
        
    Do While lPos < Len(Sql)
        
        sTmp = Mid(Sql, lPos, Len(Sql))
        
        lWhere = InStr(sTmp, UCase$(gsFindText))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            frmMain.ActiveForm.SetFocus
            frmMain.ActiveForm!txtQuery.SelStart = lPos - 2   ' set selection start and
            frmMain.ActiveForm!txtQuery.SelLength = Len(gsFindText)   ' set selection length.   Else
            gbLastPos = lPos
            Exit Do
        Else
            gbLastPos = 0
            Exit Do 'we are ready
        End If
    Loop
    
    Exit Sub
    
SalirFindText:
    gbLastPos = 0
    Err = 0
    
End Sub
Sub FileNew(Optional Index)
    
    Dim fIndex As Integer
    Dim i As Integer
    
    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    
    fState(fIndex).Tag = "q" & fIndex
    fState(fIndex).Dirty = fIndex
    Document(fIndex).Caption = "Query : " & fIndex
    Document(fIndex).Tag = "q" & fIndex
    Document(fIndex).Show
        
    Call CargaConexiones
    
    Exit Sub
    
    Dim f As Form
    Dim ci As ComboItem
    
    Dim ArrayCount As Integer
    Dim ArrayDoc As Integer

    ArrayCount = UBound(DBConnection)
    ArrayDoc = UBound(Document)
    
    Set f = Document(fIndex)
    f.imgConexiones.ComboItems.Clear
    
    For i = 1 To ArrayCount
        If Not cState(i).Deleted Then
            
            Set ci = f.imgConexiones.ComboItems.Add(1, cState(i).Conexion, _
            cState(i).Conexion, 2, 2, 0)

            ci.Indentation = 2
            
        End If
    Next i
    
    Index = fIndex
    
End Sub

Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = 1 To ArrayCount
        If fState(i).Deleted Then
            FindFreeIndex = i
            fState(i).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve fState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

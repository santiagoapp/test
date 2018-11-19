VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'
'                                   1. VARIABLES
'
'------------------------------------------------------------------------------------'

Private pTabla As String
Private pQry As String
Private pColumns As Variant
Private pValues As Variant
Private pColsFilters As Variant
Private pLogicOperators As Variant
Private pColsValues As Variant
Private pSoftDeletes As Variant
Private pJoins As Variant

Private pIfFields As Boolean
Private pIfFilters As Boolean
Private pIfSoftDeletes As Boolean
Private pIfJoins As Boolean

'------------------------------------------------------------------------------------'
'
'                                   2. PROPIEDADES
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD DE LA TABLA
'
'------------------------------------------------------------------------------------'

Property Get tabla() As String
    tabla = pTabla
End Property

Property Let tabla(value As String)
    pTabla = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.2 PROPIEDAD DE LA QUERY
'
'------------------------------------------------------------------------------------'

Property Get qry() As String
    qry = pQry
End Property

Property Let qry(value As String)
    pQry = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD LAS COLUMNAS
'
'------------------------------------------------------------------------------------'

Property Get columns() As Variant
    tabla = pColumns
End Property

Property Let columns(value As Variant)
    pColumns = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.4 PROPIEDAD LOS VALORES
'
'------------------------------------------------------------------------------------'

Property Get values() As Variant
    values = pValues
End Property

Property Let values(value As Variant)
    pValues = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.5 PROPIEDAD LAS COLUMNAS DE LOS FILTROS
'
'------------------------------------------------------------------------------------'

Property Get colsFilters() As Variant
    colsFilters = pColsFilters
End Property

Property Let colsFilters(value As Variant)
    pColsFilters = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.7 PROPIEDAD DEL OPERADOR LOGICO DE LOS FILTROS
'
'------------------------------------------------------------------------------------'

Property Get logicOperators() As Variant
    logicOperators = pLogicOperators
End Property

Property Let logicOperators(value As Variant)
    pLogicOperators = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.8 PROPIEDAD DEL VALOR DE LOS FILTROS
'
'------------------------------------------------------------------------------------'

Property Get colsValues() As Variant
    colsValues = pColsValues
End Property

Property Let colsValues(value As Variant)
    pColsValues = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.8 PROPIEDAD DE LOS SOFTDELETES
'
'------------------------------------------------------------------------------------'

Property Get softDeletes() As Variant
    softDeletes = pSoftDeletes
End Property

Property Let softDeletes(value As Variant)
    pSoftDeletes = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.9 PROPIEDAD DE LOS JOINS
'
'------------------------------------------------------------------------------------'

Property Get joins() As Variant
    joins = pJoins
End Property

Property Let joins(value As Variant)
    pJoins = value
End Property

'------------------------------------------------------------------------------------'
'
'                                   4. MÉTODOS
'
'------------------------------------------------------------------------------------'

Public Function insertQry() As String
    
    Dim strs As Variant
    
    Me.setInsertValues
    strs = buildInsertQry
    pQry = "INSERT INTO " & pTabla & " (" & strs(0) & ")" & " VALUES (" & strs(1) & ")"
    insertQry = pQry
    
End Function

Public Function buildInsertQry() As Variant
    
    Dim str(1) As String
    For col = 0 To UBound(pColumns, 1)
        str(0) = str(0) & "[" & pColumns(col) & "], "
        str(1) = str(1) & fieldByType(pValues(col)) & ", "
    Next col
    str(0) = Left(str(0), Len(str(0)) - 2)
    str(1) = Left(str(1), Len(str(1)) - 2)
    buildInsertQry = str
    
End Function

Public Function fieldByType(str As Variant) As String

    Select Case VarType(str)
        Case vbString
            If str = "NULL" Then
                fieldByType = "" & str & ""
            Else
                fieldByType = "'" & str & "'"
            End If
        Case vbInteger
            fieldByType = CStr(str)
        Case vbLong
            fieldByType = CStr(str)
        Case vbDouble
            fieldByType = CStr(str)
        Case vbBoolean
            If str Then fieldByType = "True" Else fieldByType = "False"
        Case vbDate
            fieldByType = "'" & str & "'"
        Case Else
            fieldByType = "error"
    End Select
    
End Function

Public Sub setInsertValues()
    
    Dim indice As Integer
    Dim Arr1() As Variant
    Dim Arr2() As Variant
    
    ReDim Arr1(100)
    ReDim Arr2(100)
    For j = 0 To UBound(pValues, 1)
        If VarType(pValues(j)) <> vbEmpty Then
            Arr1(indice) = pColumns(j)
            Arr2(indice) = pValues(j)
            indice = indice + 1
        End If
    Next j
    
    ReDim Preserve Arr1(indice - 1)
    ReDim Preserve Arr2(indice - 1)
    pColumns = Arr1
    pValues = Arr2
    
End Sub

Public Function showQry() As String
    
    Dim strs As Variant
    
    pIfFields = Me.ifFields
    pIfFilters = Me.ifFilters
    pIfJoins = Me.ifJoins
    pIfSoftDeletes = Me.ifSoftDeletes
    showQry = pQry
    
End Function

Public Function buildShowQry() As String
    
    Dim str As String
    For col = 0 To UBound(pColumns, 1)
        'str = str & "[" & pColumns(col) & "], "
        str = str & setBrakets(pColumns(col)) & ", "
    Next col
    str = Left(str, Len(str) - 2)
    buildShowQry = str
    
End Function
Public Function setBrakets(str As Variant) As String

    Dim str2 As String
    splited = Split(str, ".")
    If UBound(splited, 1) > 0 Then
        For Each Item In splited
            str2 = str2 & "[" & Item & "]."
        Next Item
        str = Left(str2, Len(str2) - 1) & " "
    Else
        str = " [" & str & "]"
    End If
    setBrakets = str
    
End Function
Public Function ifFields() As Boolean

    If VarType(pColumns) <> vbEmpty Then
        strs = buildShowQry
        pQry = "SELECT " & strs & " FROM [" & pTabla & "] "
        ifFields = True
    Else
        pQry = "SELECT * FROM [" & pTabla & "] "
        ifFields = False
    End If
    
End Function
Public Function ifFilters() As Boolean
    
    If VarType(pColsFilters) <> vbEmpty And _
    VarType(pLogicOperators) <> vbEmpty And _
    VarType(pColsValues) <> vbEmpty Then
    
        If UBound(pColsFilters, 1) = UBound(pLogicOperators, 1) And _
        UBound(pColsFilters, 1) = UBound(pColsValues, 1) Then
            
            For i = 0 To UBound(pColsFilters, 1)
                If i = 0 Then
                    pQry = pQry & " WHERE [" & pTabla & "].[" & pColsFilters(i) & "] " & pLogicOperators(i) & " " & fieldByType(pColsValues(i))
                Else
                    pQry = pQry & " AND [" & pTabla & "].[" & pColsFilters(i) & "] " & pLogicOperators(i) & " " & fieldByType(pColsValues(i))
                End If
            Next i
            ifFilters = True
        Else
            ifFilters = False
        End If
    Else
        ifFilters = False
    End If
End Function
Public Function ifJoins() As Boolean
    
    If VarType(pJoins) <> vbEmpty Then
        pQry = Me.buildJoinQry
        ifJoins = True
    Else
        ifJoins = False
    End If
    
End Function
Public Function ifSoftDeletes() As Boolean
    
    If pSoftDeletes And pIfFilters Then
        pQry = pQry & " AND [" & pTabla & "].[deleted_at] IS NULL"
    ElseIf pSoftDeletes And pIfFilters = False Then
        pQry = pQry & " WHERE [" & pTabla & "].[deleted_at] IS NULL"
    End If
    
End Function

Public Function updateQry(key As String, value As String) As String
    
    Dim str As String
    Me.setInsertValues
    str = buildUpdateQry
    pQry = "UPDATE [" & pTabla & "] SET " & str & " WHERE [" & pTabla & "].[" & key & "] = " & value
    updateQry = pQry
    
End Function

Public Function buildUpdateQry() As String
    
    Dim str As String
    For col = 0 To UBound(pColumns, 1)
        str = str & "[" & pColumns(col) & "] = " & fieldByType(pValues(col)) & ", "
    Next col
    str = Left(str, Len(str) - 2)
    buildUpdateQry = str
    
End Function

Public Function deleteQry(key As String, value As String) As String
    
    Dim str As String
    
    If pSoftDeletes Then
        Me.setInsertValues
        str = buildUpdateQry
        pQry = "UPDATE [" & pTabla & "] SET " & str & " WHERE [" & pTabla & "].[" & key & "] = " & value
    Else
        pQry = "DELETE FROM [" & pTabla & "] WHERE [" & pTabla & "].[" & key & "] = " & value
    End If
    deleteQry = pQry
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.3 INNER JOIN
'
'------------------------------------------------------------------------------------'

Public Function buildJoinQry() As String
    
    Dim parts(1) As String
    splited = Split(pQry, " [" & pTabla & "] ")

    For Each args In pJoins
        parts(0) = parts(0) & "("
        parts(1) = parts(1) & args(2) & " JOIN [" & args(0)(1) & "] ON [" & args(0)(0) & "].[" & args(1)(0) & "] = [" & args(0)(1) & "].[" & args(1)(1) & "]) "
    Next args
    
    buildJoinQry = splited(0) & " " & parts(0) & " [" & pTabla & "] " & parts(1) & splited(1)

End Function


'------------------------------------------------------------------------------------'
'
'   4.3
'
'------------------------------------------------------------------------------------'

Public Function getLastRecord() As String
    
    getLastRecord = "SELECT TOP 1 [id] FROM [" & pTabla & "] ORDER BY [id] DESC"

End Function
'------------------------------------------------------------------------------------'
'
'   4.3
'
'------------------------------------------------------------------------------------'

Public Function getRecordsCount() As String
    
    getRecordsCount = "SELECT COUNT(id) FROM [" & pTabla & "]"

End Function
'------------------------------------------------------------------------------------'
'
'   4.3
'
'------------------------------------------------------------------------------------'

Public Function getRecordSum(col As String, mes As Integer, año As Integer) As String
    
    getRecordSum = "SELECT SUM(" & col & ") FROM [" & pTabla & "] WHERE FORMAT(fecha_inicio,'MM') = " & Format(mes, "00") & " AND FORMAT(fecha_inicio, 'YYYY') = " & año

End Function

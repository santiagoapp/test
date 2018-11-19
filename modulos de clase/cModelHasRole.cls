VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cModelHasRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tabla As String
Private columns As Variant
Private values As Variant

Private pID As Variant
Private pRoleID As Variant
Private pUserID As Variant

Private softDeletes As Boolean
Private Con As cConnection
Private QueryBuilder As cQueryBuilder


Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property


Property Get roleID() As Variant
    roleID = pRoleID
End Property
Property Let roleID(value As Variant)
    pRoleID = value
End Property


Property Get userID() As Variant
    userID = pUserID
End Property
Property Let userID(value As Variant)
    pUserID = value
End Property

'------------------------------------------------------------------------------------'
'
'                                   3. CONSTRUCTOR
'
'------------------------------------------------------------------------------------'

Private Sub Class_Initialize()
    
    Set Con = New cConnection
    Set QueryBuilder = New cQueryBuilder
    
    softDeletes = False
    Con.DBPath = ThisWorkbook.Path & "\db\db.accdb"
    columns = Array("id", "role_id", "user_id")
    tabla = "model_has_roles"
    QueryBuilder.tabla = tabla
    QueryBuilder.softDeletes = softDeletes
    
End Sub

'------------------------------------------------------------------------------------'
'
'                                   4. MÉTODOS
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   4.1 CREAR NUEVO
'
'------------------------------------------------------------------------------------'

Public Function create() As Boolean
    
    values = Array(pID, pRoleID, pUserID)
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.insertQry
    If Con.setRecords Then create = True Else create = False
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 MOSTRAR
'
'------------------------------------------------------------------------------------'

Public Function show( _
    Optional fields As Variant = Empty, _
    Optional colsFilter As Variant = Empty, _
    Optional logicOperators As Variant = Empty, _
    Optional colsValues As Variant = Empty, _
    Optional joins As Variant = Empty _
) As Variant
    
    QueryBuilder.columns = fields
    QueryBuilder.colsFilters = colsFilter
    QueryBuilder.logicOperators = logicOperators
    QueryBuilder.colsValues = colsValues
    QueryBuilder.joins = joins
    
    Con.qry = QueryBuilder.showQry
    show = Con.getRecords
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 ACTUALIZAR
'
'------------------------------------------------------------------------------------'

Public Function update(key As String) As Boolean
    
    values = Array(pRoleID, pUserID)
    columns = Array("role_id", "user_id")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.updateQry("user_id", key)
    update = Con.setRecords
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 ELIMINAR
'
'------------------------------------------------------------------------------------'

Public Function delete(key As String) As Boolean
    
    values = Array(Empty, Empty, Empty)
    columns = Array("id", "role_id", "user_id")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("user_id", key)
    delete = Con.setRecords
    
End Function

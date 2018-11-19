VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tabla As String
Private columns As Variant
Private values As Variant

Private pID As Variant
Private pName As Variant
Private pUser As Variant
Private pPassword As Variant
Private pEmail As Variant
Private pCreatedAt As Variant
Private pUpdatedAt As Variant
Private pDeletedAt As Variant

Private softDeletes As Boolean
Private Con As cConnection
Private QueryBuilder As cQueryBuilder

Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property


Property Get name() As Variant
    name = pName
End Property
Property Let name(value As Variant)
    pName = value
End Property


Property Get User() As Variant
    User = pUser
End Property
Property Let User(value As Variant)
    pUser = value
End Property


Property Get password() As Variant
    password = pPassword
End Property
Property Let password(value As Variant)
    pPassword = value
End Property


Property Get email() As Variant
    email = pEmail
End Property
Property Let email(value As Variant)
    pEmail = value
End Property

'------------------------------------------------------------------------------------'
'
'                                   3. CONSTRUCTOR
'
'------------------------------------------------------------------------------------'

Private Sub Class_Initialize()
    
    Set Con = New cConnection
    Set QueryBuilder = New cQueryBuilder
    
    softDeletes = True
    Con.DBPath = ThisWorkbook.Path & "\db\db.accdb"
    columns = Array("id", "name", "user", "password", "email", "created_at", "updated_at", "deleted_at")
    tabla = "users"
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
    
    pCreatedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    pUpdatedAt = pCreatedAt
    values = Array(pID, pName, pUser, pPassword, pEmail, pCreatedAt, pUpdatedAt, pDeletedAt)
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.insertQry
    If Con.setRecords Then
        Con.qry = QueryBuilder.getLastRecord
        pID = Con.getRecords(0, 0)
        create = True
    Else
        create = False
    End If
    
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
    
    pUpdatedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    values = Array(pName, pUser, pPassword, pEmail, pCreatedAt, pUpdatedAt, pDeletedAt)
    columns = Array("name", "user", "password", "email", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.updateQry("id", key)
    update = Con.setRecords
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 ELIMINAR
'
'------------------------------------------------------------------------------------'

Public Function delete(key As String) As Boolean
    
    pDeletedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    values = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, pDeletedAt)
    columns = Array("id", "name", "user", "password", "email", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("id", key)
    delete = Con.setRecords
    
End Function


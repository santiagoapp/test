VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'
'                                   1. VARIABLES
'
'------------------------------------------------------------------------------------'
Private tabla As String
Private columns As Variant
Private values As Variant

Private pID As Variant
Private pCargoID As Variant
Private pPuestoID As Variant
Private pSuperiorID As Variant
Private pName As Variant
Private pLastName As Variant
Private pTipo As Variant
Private pCreatedAt As Variant
Private pUpdatedAt As Variant
Private pDeletedAt As Variant

Private softDeletes As Boolean
Private Con As cConnection
Private QueryBuilder As cQueryBuilder

'------------------------------------------------------------------------------------'
'
'                                   2. PROPIEDADES
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get ID() As Variant
    ID = pID
End Property

Public Property Let ID(value As Variant)
    pID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get cargoID() As Variant
    cargoID = pCargoID
End Property

Public Property Let cargoID(value As Variant)
    pCargoID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get puestoID() As Variant
    puestoID = pPuestoID
End Property

Public Property Let puestoID(value As Variant)
    pPuestoID = value
End Property
'------------------------------------------------------------------------------------'
'
'   2.2 PROPIEDAD DE LA RESPUESTA
'
'------------------------------------------------------------------------------------'

Public Property Get respuesta() As Variant
    respuesta = pRespuesta
End Property

Public Property Let respuesta(value As Variant)
    pRespuesta = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get superiorID() As Variant
    superiorID = pSuperiorID
End Property

Public Property Let superiorID(value As Variant)
    pSuperiorID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get name() As Variant
    name = pName
End Property

Public Property Let name(value As Variant)
    pName = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get lastName() As Variant
    lastName = pLastName
End Property

Public Property Let lastName(value As Variant)
    pLastName = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get Tipo() As Variant
    Tipo = pTipo
End Property

Public Property Let Tipo(value As Variant)
    pTipo = value
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
    columns = Array("id", "cargo_id", "puesto_id", "superior_id", "nombre", "apellido", "tipo", "created_at", "updated_at", "deleted_at")
    tabla = "personal"
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
    values = Array(pID, pCargoID, pPuestoID, pSuperiorID, pName, pLastName, pTipo, pCreatedAt, pUpdatedAt, pDeletedAt)
    
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
    
    pUpdatedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    values = Array(pCargoID, pPuestoID, pSuperiorID, pName, pLastName, pTipo, pCreatedAt, pUpdatedAt, pDeletedAt)
    columns = Array("cargo_id", "puesto_id", "superior_id", "nombre", "apellido", "tipo", "created_at", "updated_at", "deleted_at")
    
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
    values = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, pDeletedAt)
    columns = Array("id", "cargo_id", "puesto_id", "superior_id", "nombre", "apellido", "tipo", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("id", key)
    delete = Con.setRecords
    
End Function







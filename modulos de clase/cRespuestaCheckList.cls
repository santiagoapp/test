VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRespuestaChecklist"
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
Private pPreguntaID As Variant
Private pPreventivoID As Variant
Private pRespuesta As Variant
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
'   2.1 PROPIEDAD DEL ID DEL PUESTO DE TRABAJO
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
'   2.1 PROPIEDAD DEL ID DE LA PREGUNTA
'
'------------------------------------------------------------------------------------'

Public Property Get preguntaID() As Variant
    preguntaID = pPreguntaID
End Property

Public Property Let preguntaID(value As Variant)
    pPreguntaID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD DEL ID DEL MTO PREVENTIVO
'
'------------------------------------------------------------------------------------'

Public Property Get preventivoID() As Variant
    preventivoID = pPreventivoID
End Property

Public Property Let preventivoID(value As Variant)
    pPreventivoID = value
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
'   2.3 PROPIEDAD DEL ALIAS DE LA PLANTA
'
'------------------------------------------------------------------------------------'

Public Property Get Alias() As Variant
    Alias = pAlias
End Property

Public Property Let Alias(value As Variant)
    pAlias = value
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
    columns = Array("id", "pregunta_id", "preventivo_id", "respuesta", "created_at", "updated_at", "deleted_at")
    tabla = "respuestas_checklist"
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
    values = Array(pID, pPreguntaID, pPreventivoID, pRespuesta, pCreatedAt, pUpdatedAt, pDeletedAt)
    
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
    values = Array(pPreguntaID, pPreventivoID, pRespuesta, pCreatedAt, pUpdatedAt, pDeletedAt)
    columns = Array("pregunta_id", "preventivo_id", "respuesta", "created_at", "updated_at", "deleted_at")
    
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
    values = Array(Empty, Empty, Empty, Empty, Empty, Empty, pDeletedAt)
    columns = Array("id", "pregunta_id", "preventivo_id", "respuesta", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("id", key)
    delete = Con.setRecords
    
End Function





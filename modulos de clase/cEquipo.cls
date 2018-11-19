VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEquipo"
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
Private pOperarioID As Variant
Private pPuestoDeTrabajoID As Variant
Private pTipoDeMaquinaID As Variant
Private pConsecutivo As Variant
Private pNombreEquipo As Variant
Private pTipoDeRevision As Variant
Private pDescripcion As Variant
Private pObservacion As Variant
Private pAlias As Variant
Private pCosto As Variant
Private pEstado As Variant
Private pMarca As Variant
Private pModelo As Variant
Private pNumeroFactura As Variant
Private pCorriente As Variant
Private pVoltaje As Variant
Private pProcedencia As Variant
Private pCatalogo As Variant
Private pDibujo As Variant
Private pDiagramaElectrico As Variant
Private pManual As Variant
Private pNumeroDeSerie As Variant
Private pFechaCompra As Variant
Private pFechaPrimerUso As Variant
Private pImportancia As Variant
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

Public Property Get operarioID() As Variant
    operarioID = pOperarioID
End Property

Public Property Let operarioID(value As Variant)
    pOperarioID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get puestoDeTrabajoID() As Variant
    puestoDeTrabajoID = pPuestoDeTrabajoID
End Property

Public Property Let puestoDeTrabajoID(value As Variant)
    pPuestoDeTrabajoID = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get tipoDeMaquinaID() As Variant
    tipoDeMaquinaID = pTipoDeMaquinaID
End Property

Public Property Let tipoDeMaquinaID(value As Variant)
    pTipoDeMaquinaID = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get consecutivo() As Variant
    consecutivo = pConsecutivo
End Property

Public Property Let consecutivo(value As Variant)
    pConsecutivo = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get nombreEquipo() As Variant
    nombreEquipo = pNombreEquipo
End Property

Public Property Let nombreEquipo(value As Variant)
    pNombreEquipo = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get tipoDeRevision() As Variant
    tipoDeRevision = pTipoDeRevision
End Property

Public Property Let tipoDeRevision(value As Variant)
    pTipoDeRevision = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get Descripcion() As Variant
    Descripcion = pDescripcion
End Property

Public Property Let Descripcion(value As Variant)
    pDescripcion = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get observacion() As Variant
    observacion = pObservacion
End Property

Public Property Let observacion(value As Variant)
    pObservacion = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
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
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get Costo() As Variant
    Costo = pCosto
End Property

Public Property Let Costo(value As Variant)
    pCosto = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get estado() As Variant
    estado = pEstado
End Property

Public Property Let estado(value As Variant)
    pEstado = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get Marca() As Variant
    Marca = pMarca
End Property

Public Property Let Marca(value As Variant)
    pMarca = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get modelo() As Variant
    modelo = pModelo
End Property

Public Property Let modelo(value As Variant)
    pModelo = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get numeroFactura() As Variant
    numeroFactura = pNumeroFactura
End Property

Public Property Let numeroFactura(value As Variant)
    pNumeroFactura = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get corriente() As Variant
    corriente = pCorriente
End Property

Public Property Let corriente(value As Variant)
    pCorriente = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get voltaje() As Variant
    voltaje = pVoltaje
End Property

Public Property Let voltaje(value As Variant)
    pVoltaje = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get procedencia() As Variant
    procedencia = pProcedencia
End Property

Public Property Let procedencia(value As Variant)
    pProcedencia = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get catalogo() As Variant
    catalogo = pCatalogo
End Property

Public Property Let catalogo(value As Variant)
    pCatalogo = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get diagramaElectrico() As Variant
    diagramaElectrico = pDiagramaElectrico
End Property

Public Property Let diagramaElectrico(value As Variant)
    pDiagramaElectrico = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get dibujo() As Variant
    dibujo = pDibujo
End Property

Public Property Let dibujo(value As Variant)
    pDibujo = value
End Property
'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get manual() As Variant
    manual = pManual
End Property

Public Property Let manual(value As Variant)
    pManual = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get numeroDeSerie() As Variant
    numeroDeSerie = pNumeroDeSerie
End Property

Public Property Let numeroDeSerie(value As Variant)
    pNumeroDeSerie = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get fechaCompra() As Variant
    fechaCompra = pFechaCompra
End Property

Public Property Let fechaCompra(value As Variant)
    pFechaCompra = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get fechaPrimerUso() As Variant
    fechaPrimerUso = pFechaPrimerUso
End Property

Public Property Let fechaPrimerUso(value As Variant)
    pFechaPrimerUso = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get importancia() As Variant
    importancia = pImportancia
End Property

Public Property Let importancia(value As Variant)
    pImportancia = value
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
    columns = Array("id", "operario_id", "puesto_de_trabajo_id", "tipo_de_maquina_id", "consecutivo", "nombre_equipo", "tipo_de_revision", "descripcion", "observacion", "alias", "costo", "estado", "marca", "modelo", "numero_factura", "corriente", "voltaje", "procedencia", "catalogo", "dibujo", "diagrama_electrico", "manual", "numero_de_serie", "fecha_compra", "fecha_primer_uso", "importancia", "created_at", "updated_at", "deleted_at")
    tabla = "equipos"
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
    values = Array(pID, pOperarioID, pPuestoDeTrabajoID, pTipoDeMaquinaID, pConsecutivo, pNombreEquipo, pTipoDeRevision, pDescripcion, pObservacion, pAlias, pCosto, pEstado, pMarca, pModelo, pNumeroFactura, pCorriente, pVoltaje, pProcedencia, pCatalogo, pDibujo, pDiagramaElectrico, pManual, pNumeroDeSerie, pFechaCompra, pFechaPrimerUso, pImportancia, pCreatedAt, pUpdatedAt, pDeletedAt)

    
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
    values = Array(pOperarioID, pPuestoDeTrabajoID, pTipoDeMaquinaID, pConsecutivo, pNombreEquipo, pTipoDeRevision, pDescripcion, pObservacion, pAlias, pCosto, pEstado, pMarca, pModelo, pNumeroFactura, pCorriente, pVoltaje, pProcedencia, pCatalogo, pDibujo, pDiagramaElectrico, pManual, pNumeroDeSerie, pFechaCompra, pFechaPrimerUso, pImportancia, pCreatedAt, pUpdatedAt, pDeletedAt)
    columns = Array("operario_id", "puesto_de_trabajo_id", "tipo_de_maquina_id", "consecutivo", "nombre_equipo", "tipo_de_revision", "descripcion", "observacion", "alias", "costo", "estado", "marca", "modelo", "numero_factura", "corriente", "voltaje", "procedencia", "catalogo", "dibujo", "diagrama_electrico", "manual", "numero_de_serie", "fecha_compra", "fecha_primer_uso", "importancia", "created_at", "updated_at", "deleted_at")
    
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
    values = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, pDeletedAt)
    columns = Array("id", "operario_id", "puesto_de_trabajo_id", "tipo_de_maquina_id", "consecutivo", "nombre_equipo", "tipo_de_revision", "descripcion", "observacion", "alias", "costo", "estado", "marca", "modelo", "numero_factura", "corriente", "voltaje", "procedencia", "catalogo", "dibujo", "diagrama_electrico", "manual", "numero_de_serie", "fecha_compra", "fecha_primer_uso", "importancia", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("id", key)
    delete = Con.setRecords
    
End Function

Public Function setTemporalID() As Integer
    
    Con.qry = QueryBuilder.getLastRecord
    setTemporalID = Con.getRecords(0, 0) + 1
    
End Function

Public Function count() As Integer
    
    Con.qry = QueryBuilder.getRecordsCount
    count = Con.getRecords(0, 0)
    
End Function









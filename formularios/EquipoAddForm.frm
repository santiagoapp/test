VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EquipoAddForm 
   Caption         =   "Agregar nuevo activo"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "EquipoAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "EquipoAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Form As cForm
Private Equipo As cEquipo
Private Personal As cPersonal
Private Planta As cPlanta
Private PuestoDeTrabajo As cPuestoDeTrabajo
Private TipoDeMaquina As cTipoDeMaquina

Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub CommandButton3_Click()

    Equipo.ID = pID
    Equipo.nombreEquipo = Me.NombreField
    Equipo.puestoDeTrabajoID = Me.PuestoDeTrabajoField
    Equipo.operarioID = Me.ResponsableField
    Equipo.Descripcion = Me.DescripcionField
    Equipo.observacion = Me.ObservacionField
    Equipo.tipoDeMaquinaID = Me.TipoDeSistemaField
    Equipo.Marca = Me.MarcaField
    Equipo.modelo = Me.ModeloField
    Equipo.numeroDeSerie = Me.NumeroDeSerieField
    Equipo.numeroFactura = Me.NumeroDeFacturaField
    Equipo.corriente = Me.TipoDeCorrienteField
    Equipo.voltaje = Me.VoltajeField
    Equipo.procedencia = Me.PaisField
    Equipo.fechaCompra = Me.FechaCompraField
    Equipo.fechaPrimerUso = Me.FechaUsoField
    Equipo.catalogo = Me.CatalogoField
    Equipo.diagramaElectrico = Me.DiagramaField
    Equipo.dibujo = Me.DibujoField
    Equipo.manual = Me.ManualField
    Equipo.Costo = Me.CostoField
    Equipo.Alias = Me.AliasField
    Equipo.importancia = Me.ImportanciaField
    Equipo.tipoDeRevision = Me.TipoDeRevisionField
    Equipo.estado = Me.EstadoField
    Equipo.consecutivo = Me.ConsecutivoField
    
    If VarType(pID) = vbEmpty Then Call Equipo.create Else Call Equipo.update(CInt(pID))
    Me.Hide
    
End Sub


Private Sub Image2_Click()
    FechaUsoField.value = Form.Calendar
End Sub

Private Sub img_Click()
    FechaCompraField.value = Form.Calendar
End Sub

Private Sub PlantaField_Change()

    If VarType(Me.PlantaField) <> vbNull Then
        Set PuestoDeTrabajo = New cPuestoDeTrabajo
        puestos = PuestoDeTrabajo.show( _
            fields:=Array("id", "nombre"), _
            colsFilter:=Array("planta_id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.PlantaField)) _
        )
        Call Form.fillListOrComboBox(puestos, Me.PuestoDeTrabajoField)
    End If
    
End Sub


Private Sub UserForm_Initialize()
    
    Set Form = New cForm
    Set Equipo = New cEquipo
    Set Personal = New cPersonal
    Set Planta = New cPlanta
    Set TipoDeMaquina = New cTipoDeMaquina

    Me.EstadoField.AddItem "Activo"
    Me.EstadoField.AddItem "Dañado"
    
    Me.TipoDeRevisionField.AddItem "Baja"
    Me.TipoDeRevisionField.AddItem "Normal"
    Me.TipoDeRevisionField.AddItem "Exhaustiva"
    
    For i = 1 To 5
        Me.ImportanciaField.AddItem i
    Next i
    personas = Personal.show( _
        fields:=Array("id", "nombre"), _
        colsFilter:=Array("tipo"), _
        logicOperators:=Array("="), _
        colsValues:=Array("Operativo") _
    )
    plantas = Planta.show( _
        fields:=Array("id", "nombre") _
    )
    tipos = TipoDeMaquina.show( _
        fields:=Array("id", "tipo_maquina") _
    )
    
    
    Call Form.fillListOrComboBox(personas, Me.ResponsableField)
    Call Form.fillListOrComboBox(plantas, Me.PlantaField)
    Call Form.fillListOrComboBox(tipos, Me.TipoDeSistemaField)
    
    
End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EquipoMainForm 
   Caption         =   "Administración de los equipos"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "EquipoMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "EquipoMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Equipo As cEquipo
Private Form As cForm

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("EquipoAddForm")
    AgregarForm.ConsecutivoField = "FL" & Format(Equipo.setTemporalID, "000")
    AgregarForm.show
    Unload AgregarForm
    Set AgregarForm = Nothing
    Call Me.getFields
    
End Sub

Private Sub borrar_Click()
    
    mensaje = MsgBox("¿Desea continuar con la eliminación del registro?", vbInformation + vbYesNo)
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        If mensaje = vbYes Then
            response = Equipo.delete(CInt(Me.ListBox1))
            If response Then MsgBox "Registro eliminado con éxito", vbInformation
        End If
    End If
    Call Me.getFields
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub correctivoBtn_Click()

    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set AgregarForm = UserForms.Add("CorrectivoAddForm")
        AgregarForm.equipoID = Me.ListBox1
        AgregarForm.show
        Unload AgregarForm
        Set AgregarForm = Nothing
        Call Me.getFields
    End If
    
End Sub


Private Sub ListBox1_Click()
    
    If VarType(Me.ListBox1) <> vbNull Then
        Set Equipo = New cEquipo

        joins = Array( _
                Array( _
                    Array("equipos", "personal"), _
                    Array("operario_id", "id"), _
                    "LEFT"), _
                Array( _
                    Array("equipos", "puestos_de_trabajo"), _
                    Array("puesto_de_trabajo_id", "id"), _
                    "LEFT"), _
                Array( _
                    Array("equipos", "tipos_de_maquinas"), _
                    Array("tipo_de_maquina_id", "id"), _
                    "LEFT") _
            )
            
        arr = Equipo.show( _
            fields:=Array("equipos.nombre_equipo", "personal.nombre", "equipos.costo", "puestos_de_trabajo.nombre", "equipos.marca", "tipos_de_maquinas.tipo_maquina", "equipos.descripcion", "equipos.observacion"), _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)), _
            joins:=joins _
        )
        
        Me.Nombre = arr(0, 0)
        Me.Responsable = arr(1, 0)
        Me.Costo = Format(CDbl(arr(2, 0)), "$#,##0.00;-$#,##0.00")
        Me.Zona = arr(3, 0)
        Me.Marca = arr(4, 0)
        Me.TipoDeMaquina = arr(5, 0)
        Me.Descripcion = arr(6, 0)
        Me.Observaciones = arr(7, 0)
        
    End If
    
End Sub

Private Sub modificar_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub modificar_Click()
    
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set EditarForm = UserForms.Add("EquipoAddForm")
        Set Equipo = New cEquipo
        EditarForm.ID = Me.ListBox1
        
        joins = Array( _
            Array( _
                Array("equipos", "puestos_de_trabajo"), _
                Array("puesto_de_trabajo_id", "id"), _
                "INNER" _
            ), _
            Array( _
                Array("puestos_de_trabajo", "plantas"), _
                Array("planta_id", "id"), _
                "INNER" _
            ) _
        )
        
        arr = Equipo.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)), _
            joins:=joins _
        )
        
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar máquina"
        
        EditarForm.NombreField = arr(5, 0)
        EditarForm.PlantaField = arr(30, 0)
        EditarForm.PuestoDeTrabajoField = arr(2, 0)
        EditarForm.ResponsableField = arr(1, 0)
        EditarForm.DescripcionField = arr(7, 0)
        EditarForm.ObservacionField = arr(8, 0)
        EditarForm.TipoDeSistemaField = arr(3, 0)
        EditarForm.MarcaField = arr(12, 0)
        EditarForm.ModeloField = arr(13, 0)
        EditarForm.NumeroDeSerieField = arr(22, 0)
        EditarForm.NumeroDeFacturaField = arr(14, 0)
        EditarForm.TipoDeCorrienteField = arr(15, 0)
        EditarForm.VoltajeField = arr(16, 0)
        EditarForm.PaisField = arr(17, 0)
        EditarForm.FechaCompraField = arr(23, 0)
        EditarForm.FechaUsoField = arr(24, 0)
        EditarForm.CatalogoField = arr(18, 0)
        EditarForm.DiagramaField = arr(20, 0)
        EditarForm.DibujoField = arr(19, 0)
        EditarForm.ManualField = arr(21, 0)
        EditarForm.CostoField = arr(10, 0)
        EditarForm.AliasField = arr(9, 0)
        EditarForm.ImportanciaField = arr(25, 0)
        EditarForm.TipoDeRevisionField = arr(6, 0)
        EditarForm.EstadoField = arr(11, 0)
        EditarForm.ConsecutivoField = arr(4, 0)
        
        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set Equipo = New cEquipo
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()
    joins = Array( _
                Array( _
                    Array("equipos", "personal"), _
                    Array("operario_id", "id"), _
                    "LEFT"), _
                Array( _
                    Array("equipos", "puestos_de_trabajo"), _
                    Array("puesto_de_trabajo_id", "id"), _
                    "LEFT") _
            )
    arr = Equipo.show( _
        fields:=Array("equipos.id", "equipos.nombre_equipo", "personal.nombre", "equipos.estado", "puestos_de_trabajo.nombre", "equipos.importancia"), _
        joins:=joins _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function






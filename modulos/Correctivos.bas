Attribute VB_Name = "Correctivos"
Private Correctivo As cCorrectivo
Private Form As cForm

Public Sub completarCorrectivo()
    
    Dim fecha_fin As Date
    Dim fecha_inicio As Date
    
    If isRowSelected(Hoja8.ListBox1) Then Exit Sub
        
    response = MsgBox("¿Está seguro de desea completar el mantenimiento correctivo?", vbYesNo + vbInformation)
    If response = vbYes Then
        
        Set Correctivo = New cCorrectivo
        Set Form = New cForm
        
        fecha_fin = Form.Calendar
        fecha_inicio = Hoja8.ListBox1.List(, 6)
        If Day(fecha_inicio) < Day(fecha_fin) And _
            Month(fecha_inicio) <= Month(fecha_fin) And _
            year(fecha_inicio) <= year(fecha_fin) _
        Then
            Correctivo.fechaFin = fecha_fin
            Correctivo.update (CInt(Hoja8.ListBox1))
            Call Correctivos.getCorrectivosActivos
        Else
            MsgBox "La fecha es inferior a la fecha de inicio, por favor seleccione una fecha correcta", vbInformation + vbOKOnly
        End If
        
    End If
    
    
End Sub

Public Sub getCorrectivosActivos()
    
    Set Correctivo = New cCorrectivo
    Set Form = New cForm
    
    joins = Array( _
        Array( _
            Array("correctivos", "equipos"), _
            Array("equipo_id", "id"), _
            "INNER" _
        ), _
        Array( _
            Array("correctivos", "personal"), _
            Array("encargado_id", "id"), _
            "INNER" _
        ), _
        Array( _
            Array("equipos", "puestos_de_trabajo"), _
            Array("puesto_de_trabajo_id", "id"), _
            "INNER" _
        ) _
    ) _

    fields = Array( _
        "correctivos.id", _
        "equipos.id", _
        "equipos.consecutivo", _
        "equipos.nombre_equipo", _
        "puestos_de_trabajo.nombre", _
        "personal.nombre", _
        "correctivos.fecha_inicio" _
    )
    
    correctivosActivos = Correctivo.show( _
        fields:=fields, _
        joins:=joins, _
        colsFilter:=Array("fecha_fin"), _
        logicOperators:=Array("IS"), _
        colsValues:=Array("NULL") _
    )
    
    Call Form.fillListOrComboBox(correctivosActivos, Hoja8.ListBox1)
    
End Sub
Public Function isRowSelected(ListBox As Object) As Boolean
    If VarType(ListBox) = vbNull Then
        MsgBox "Seleccione un registro para continuar", vbOKOnly + vbInformation
        isRowSelected = True
    Else
        isRowSelected = False
    End If
End Function

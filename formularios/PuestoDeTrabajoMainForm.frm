VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PuestoDeTrabajoMainForm 
   Caption         =   "Administración de los puestos de trabajo"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "PuestoDeTrabajoMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PuestoDeTrabajoMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PuestoDeTrabajo As cPuestoDeTrabajo
Private Form As cForm

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("PuestoDeTrabajoAddForm")
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
            response = PuestoDeTrabajo.delete(CInt(Me.ListBox1))
            If response Then MsgBox "Registro eliminado con éxito", vbInformation
        End If
    End If
    Call Me.getFields
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub modificar_Click()
    
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set EditarForm = UserForms.Add("PuestoDeTrabajoAddForm")
        Set PuestoDeTrabajo = New cPuestoDeTrabajo
        EditarForm.ID = Me.ListBox1
        arr = PuestoDeTrabajo.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar puesto de trabajo"
        
        EditarForm.Nombre = arr(2, 0)
        EditarForm.Alias = arr(3, 0)
        EditarForm.ComboBox3 = arr(1, 0)
        
        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set PuestoDeTrabajo = New cPuestoDeTrabajo
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()
    
    joins = Array( _
                Array( _
                    Array("puestos_de_trabajo", "plantas"), _
                    Array("planta_id", "id"), _
                    "INNER") _
            )
    arr = PuestoDeTrabajo.show( _
        fields:=Array("puestos_de_trabajo.id", "puestos_de_trabajo.nombre", "puestos_de_trabajo.alias", "plantas.nombre"), _
        joins:=joins _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function



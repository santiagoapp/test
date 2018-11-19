VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RoleMainForm 
   Caption         =   "Administración de roles"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "RoleMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RoleMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private role As cRole
Private PreguntaCheckList As cPreguntaCheckList
Private Form As cForm

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("RoleAddForm")
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
            response = role.delete(CInt(Me.ListBox1))
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
        Set EditarForm = UserForms.Add("RoleAddForm")
        Set role = New cRole
        EditarForm.ID = Me.ListBox1
        arr = role.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar rol"
        
        EditarForm.ID = arr(0, 0)
        EditarForm.NameField = arr(1, 0)

        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set role = New cRole
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()

    roles = role.show( _
        fields:=Array("id", "name") _
    )
    Call Form.fillListOrComboBox(roles, Me.ListBox1)
    
End Function


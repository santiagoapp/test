VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserMainForm 
   Caption         =   "Administración de usuarios"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   OleObjectBlob   =   "UserMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private User As cUser
Private Form As cForm
Private joins As Variant

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("UserAddForm")
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
            response = User.delete(CInt(Me.ListBox1))
            If response Then MsgBox "Registro eliminado con éxito", vbInformation
        End If
    End If
    Call Me.getFields
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub modificar_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub modificar_Click()
    
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set EditarForm = UserForms.Add("UserAddForm")
        Set User = New cUser
        EditarForm.ID = Me.ListBox1
        arr = User.show( _
            fields:=Array("users.id", "users.user", "users.name", "users.email", "model_has_roles.role_id", "users.password"), _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)), _
            joins:=joins _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar usuario"
        
        EditarForm.UserField = arr(1, 0)
        EditarForm.NameField = arr(2, 0)
        EditarForm.EmailField = arr(3, 0)
        EditarForm.RoleField = arr(4, 0)
        EditarForm.PasswordField1 = arr(5, 0)
        EditarForm.PasswordField2 = arr(5, 0)
        
        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set User = New cUser
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()
    
    Dim arr As Variant
    joins = Array( _
        Array( _
            Array("users", "model_has_roles"), _
            Array("id", "user_id"), _
            "INNER" _
        ), _
        Array( _
            Array("model_has_roles", "roles"), _
            Array("role_id", "id"), _
            "INNER" _
        ) _
    )
    arr = User.show( _
        fields:=Array("users.id", "users.user", "users.name", "users.email", "roles.name"), _
        joins:=joins _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function





VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PersonaMainForm 
   Caption         =   "Administración del personal"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   OleObjectBlob   =   "PersonaMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PersonaMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Personal As cPersonal
Private Form As cForm

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("PersonaAddForm")
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
            response = Personal.delete(CInt(Me.ListBox1))
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
        Set EditarForm = UserForms.Add("PersonaAddForm")
        Set Personal = New cPersonal
        EditarForm.ID = Me.ListBox1
        arr = Personal.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar funcionario"
        EditarForm.ComboBox3 = arr(1, 0)
        EditarForm.ComboBox4 = arr(2, 0)
        EditarForm.Nombre = arr(4, 0)
        EditarForm.ComboBox1 = arr(6, 0)
        If arr(3, 0) <> 0 Then EditarForm.ComboBox5 = arr(3, 0)
        EditarForm.Apellido = arr(5, 0)

        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set Personal = New cPersonal
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()

    arr = Personal.show( _
        fields:=Array("id", "nombre", "apellido", "tipo") _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function

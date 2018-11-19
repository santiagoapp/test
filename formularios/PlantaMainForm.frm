VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlantaMainForm 
   Caption         =   "Administración de facilities"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   OleObjectBlob   =   "PlantaMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PlantaMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Planta As cPlanta
Private Form As cForm

Private Sub agregar_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("PlantaAddForm")
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
            response = Planta.delete(CInt(Me.ListBox1))
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
        Set EditarForm = UserForms.Add("PlantaAddForm")
        Set Planta = New cPlanta
        EditarForm.ID = Me.ListBox1
        arr = Planta.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar planta"
        EditarForm.Nombre = arr(1, 0)
        EditarForm.Alias = arr(2, 0)

        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub UserForm_Initialize()

    Set Planta = New cPlanta
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()

    arr = Planta.show( _
        fields:=Array("id", "nombre", "alias") _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function



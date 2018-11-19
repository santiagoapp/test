VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AsignarPreguntasForm 
   Caption         =   "Preguntas para tipo de máquina"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   OleObjectBlob   =   "AsignarPreguntasForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "AsignarPreguntasForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pTipoDeMaquinaID As Variant
Private PreguntaCheckList As cPreguntaCheckList
Private Form As cForm

Property Get tipoDeMaquinaID() As Variant
    tipoDeMaquinaID = pTipoDeMaquinaID
End Property

Property Let tipoDeMaquinaID(value As Variant)
    pTipoDeMaquinaID = value
End Property

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("AgregarPreguntaForm")
    AgregarForm.tipoDeMaquinaID = pTipoDeMaquinaID
    AgregarForm.show
    Unload AgregarForm
    Set AgregarForm = Nothing
    Me.getFields
    
End Sub

Private Sub borrar_Click()
    
    mensaje = MsgBox("¿Desea continuar con la eliminación del registro?", vbInformation + vbYesNo)
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        If mensaje = vbYes Then
            response = PreguntaCheckList.delete(CInt(Me.ListBox1))
            If response Then MsgBox "Registro eliminado con éxito", vbInformation
        End If
    End If
    Call Me.getFields
    
End Sub

Private Sub modificar_Click()
    
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set EditarForm = UserForms.Add("AgregarPreguntaForm")
        Set PreguntaCheckList = New cPreguntaCheckList
        EditarForm.ID = Me.ListBox1
        arr = PreguntaCheckList.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar pregunta"
        
        EditarForm.ID = arr(0, 0)
        EditarForm.Pregunta = arr(2, 0)
        
        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub
Private Sub UserForm_Initialize()
    
    Set PreguntaCheckList = New cPreguntaCheckList
    
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Public Function getFields()

    Set Form = New cForm
    arr = PreguntaCheckList.show( _
        fields:=Array("id", "pregunta"), _
        colsFilter:=Array("tipo_de_maquina_id"), _
        logicOperators:=Array("="), _
        colsValues:=Array(pTipoDeMaquinaID) _
    )
    
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function


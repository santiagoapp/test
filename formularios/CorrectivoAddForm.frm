VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CorrectivoAddForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "CorrectivoAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CorrectivoAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private pEquipoID As Variant
Private Form As cForm
Private Correctivo As cCorrectivo
Private Encargado As cPersonal

Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property

Property Get equipoID() As Variant
    equipoID = pEquipoID
End Property
Property Let equipoID(value As Variant)
    pEquipoID = value
End Property


Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub CommandButton3_Click()

    Correctivo.ID = pID

    Correctivo.equipoID = Me.equipoID
    Correctivo.encargadoID = Me.ResponsableField
    Correctivo.tipoDeFalla = Me.FallaField
    Correctivo.causaDeFalla = Me.CausaField
    Correctivo.provedor = Me.ProvedorField
    Correctivo.Costo = Me.CostoField
    Correctivo.fechaInicio = Me.FechaInicioField
    
    If VarType(pID) = vbEmpty Then Call Correctivo.create Else Call Correctivo.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub Image2_Click()
    FechaInicioField.value = Form.Calendar
End Sub

Private Sub UserForm_Initialize()
    
    Set Form = New cForm
    Set Correctivo = New cCorrectivo
    Set Encargado = New cPersonal
    
    Me.ProvedorField.AddItem "Interno"
    Me.ProvedorField.AddItem "Externo"
    
    personas = Encargado.show( _
        fields:=Array("id", "nombre") _
    )

    Call Form.fillListOrComboBox(personas, Me.ResponsableField)

End Sub



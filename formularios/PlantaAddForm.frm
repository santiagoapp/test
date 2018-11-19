VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlantaAddForm 
   Caption         =   "Agregar planta"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "PlantaAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PlantaAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Planta As cPlanta

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

    Planta.ID = pID

    Planta.name = Me.Nombre
    Planta.Alias = Me.Alias
    
    If VarType(pID) = vbEmpty Then Call Planta.create Else Call Planta.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set Planta = New cPlanta

End Sub



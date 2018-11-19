VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CargoAddForm 
   Caption         =   "Agregar cargo"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "CargoAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CargoAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Cargo As cCargo

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

    Cargo.ID = pID

    Cargo.name = Me.Nombre
    Cargo.Descripcion = Me.Descripcion
    
    If VarType(pID) = vbEmpty Then Call Cargo.create Else Call Cargo.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set Cargo = New cCargo

End Sub


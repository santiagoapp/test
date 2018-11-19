VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PuestoDeTrabajoAddForm 
   Caption         =   "Agregar puesto de trabajo"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "PuestoDeTrabajoAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PuestoDeTrabajoAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Planta As cPlanta
Private PuestoDeTrabajo As cPuestoDeTrabajo
Private Form As New cForm

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
    
    PuestoDeTrabajo.ID = pID
    PuestoDeTrabajo.name = Me.Nombre
    PuestoDeTrabajo.Alias = Me.Alias
    PuestoDeTrabajo.plantaID = Me.ComboBox3
        
    If VarType(pID) = vbEmpty Then Call PuestoDeTrabajo.create Else Call PuestoDeTrabajo.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    
    Set Planta = New cPlanta
    Set PuestoDeTrabajo = New cPuestoDeTrabajo
    Set Form = New cForm

    plantas = Planta.show( _
        fields:=Array("id", "nombre") _
    )
    
    Call Form.fillListOrComboBox(plantas, Me.ComboBox3)
    
End Sub



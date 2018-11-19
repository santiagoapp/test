VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PersonaAddForm 
   Caption         =   "Agregar Funcionario"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "PersonaAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PersonaAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Cargo As cCargo
Private PuestoDeTrabajo As cPuestoDeTrabajo
Private Personal As cPersonal
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

    Personal.ID = pID
    Personal.cargoID = Me.ComboBox3
    Personal.puestoID = Me.ComboBox4
    Personal.name = Me.Nombre
    Personal.Tipo = Me.ComboBox1
    If Me.ComboBox5 = "" Then Personal.superiorID = Empty
    Personal.lastName = Me.Apellido
    If VarType(pID) = vbEmpty Then Call Personal.create Else Call Personal.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    
    Set Cargo = New cCargo
    Set PuestoDeTrabajo = New cPuestoDeTrabajo
    Set Personal = New cPersonal
    Set Form = New cForm
    
    Me.ComboBox1.AddItem "Administrativo"
    Me.ComboBox1.AddItem "Operativo"
    
    personas = Personal.show( _
        fields:=Array("id", "nombre") _
    )
    cargos = Cargo.show( _
        fields:=Array("id", "nombre") _
    )
    puestos = PuestoDeTrabajo.show( _
        fields:=Array("id", "nombre") _
    )
    
    Call Form.fillListOrComboBox(personas, Me.ComboBox5)
    Call Form.fillListOrComboBox(cargos, Me.ComboBox3)
    Call Form.fillListOrComboBox(puestos, Me.ComboBox4)
    
End Sub

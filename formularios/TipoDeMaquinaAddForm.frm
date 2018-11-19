VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TipoDeMaquinaAddForm 
   Caption         =   "Agregar tipo de máquina"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "TipoDeMaquinaAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TipoDeMaquinaAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private TipoDeMaquina As cTipoDeMaquina

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

    TipoDeMaquina.ID = pID
    TipoDeMaquina.Tipo = Me.Tipo
    
    If VarType(pID) = vbEmpty Then Call TipoDeMaquina.create Else Call TipoDeMaquina.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set TipoDeMaquina = New cTipoDeMaquina

End Sub



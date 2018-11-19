VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AgregarPreguntaForm 
   Caption         =   "Agregar pregunta"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "AgregarPreguntaForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "AgregarPreguntaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private PreguntaCheckList As cPreguntaCheckList
Private pTipoDeMaquinaID As Variant

Property Get tipoDeMaquinaID() As Variant
    tipoDeMaquinaID = pTipoDeMaquinaID
End Property

Property Let tipoDeMaquinaID(value As Variant)
    pTipoDeMaquinaID = value
End Property

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

    PreguntaCheckList.ID = pID
    PreguntaCheckList.tipoDeMaquinaID = pTipoDeMaquinaID
    PreguntaCheckList.Pregunta = Me.Pregunta
    
    If VarType(pID) = vbEmpty Then Call PreguntaCheckList.create Else Call PreguntaCheckList.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set PreguntaCheckList = New cPreguntaCheckList

End Sub



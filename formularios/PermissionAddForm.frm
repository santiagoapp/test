VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PermissionAddForm 
   Caption         =   "Agregar permiso"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   OleObjectBlob   =   "PermissionAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PermissionAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Permission As cPermission

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

    Permission.ID = pID
    Permission.name = Me.NameField
    
    If VarType(pID) = vbEmpty Then Call Permission.create Else Call Permission.update(CInt(pID))
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set Permission = New cPermission

End Sub




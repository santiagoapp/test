VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserAddForm 
   Caption         =   "Agregar Usuario"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "UserAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private User As cUser
Private role As cRole
Private ModelHasRole As cModelHasRole
Private Form As cForm

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

    User.ID = pID
    User.name = Me.NameField
    User.User = Me.UserField
    User.email = Me.EmailField
    If Me.PasswordField1 = Me.PasswordField2 Then User.password = Me.PasswordField1 Else MsgBox "Las contraseñas no coinciden", vbInformation + vbOKOnly: Exit Sub
    ModelHasRole.roleID = CInt(Me.RoleField)
    
    If VarType(pID) = vbEmpty Then
        Call User.create
        ModelHasRole.userID = User.ID
        Call ModelHasRole.create
    Else
        Call User.update(CInt(pID))
        Call ModelHasRole.update(CInt(pID))
    End If
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Set User = New cUser
    Set role = New cRole
    Set Form = New cForm
    Set ModelHasRole = New cModelHasRole
    
    roles = role.show( _
        fields:=Array("id", "name") _
    )
    Call Form.fillListOrComboBox(roles, Me.RoleField)
    
End Sub




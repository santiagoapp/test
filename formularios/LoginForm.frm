VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Iniciar sesión"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private User As cUser

Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property

Private Sub CommandButton2_Click()
    ThisWorkbook.Close
End Sub

Private Sub CommandButton3_Click()
    
    Dim UserCredentials As Variant
    Dim arr As Variant
    Dim i As Integer
    
    ReDim arr(1, 1000)
    joins = Array( _
        Array( _
            Array("users", "model_has_roles"), _
            Array("id", "user_id"), _
            "INNER" _
        ), _
        Array( _
            Array("model_has_roles", "roles"), _
            Array("role_id", "id"), _
            "INNER" _
        ), _
        Array( _
            Array("roles", "role_has_permissions"), _
            Array("id", "role_id"), _
            "INNER" _
        ), _
        Array( _
            Array("role_has_permissions", "permissions"), _
            Array("permission_id", "id"), _
            "INNER" _
        ) _
    )
    
    fields = Array("users.name", "users.user", "users.email", "model_has_roles.role_id", "roles.name", "permissions.id", "permissions.name")
    
    UserCredentials = User.show( _
        fields:=fields, _
        joins:=joins, _
        colsFilter:=Array("user", "password"), _
        logicOperators:=Array("=", "="), _
        colsValues:=Array(Me.UserField, Me.PasswordField1) _
    )
    If VarType(UserCredentials) = vbEmpty Then
        MsgBox "El usuario o la contraseña no se encuentran en nuestra base de datos", vbExclamation + vbOKOnly
    Else
        
        ThisWorkbook.username = UserCredentials(1, 0)
        ThisWorkbook.email = UserCredentials(2, 0)
        ThisWorkbook.roleID = UserCredentials(3, 0)
        ThisWorkbook.role = UserCredentials(4, 0)
        For i = LBound(UserCredentials, 2) To UBound(UserCredentials, 2)
            arr(0, i) = UserCredentials(5, i)
            arr(1, i) = UserCredentials(6, i)
        Next i
        ReDim Preserve arr(1, i - 1)
        ThisWorkbook.permisos = arr
        MsgBox "Bienvenido " & UserCredentials(0, 0), vbInformation + vbOKOnly
    End If
    Me.Hide
    
End Sub
Private Sub UserForm_Initialize()

    Set User = New cUser

End Sub




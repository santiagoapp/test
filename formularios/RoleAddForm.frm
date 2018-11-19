VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RoleAddForm 
   Caption         =   "Agregar Rol"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   OleObjectBlob   =   "RoleAddForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RoleAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pID As Variant
Private Permission As cPermission
Private RoleHasPermission As cRoleHasPermission
Private role As cRole
Private Form As New cForm
Private permisos As Variant
Private permisosPorRol As Variant

Property Get ID() As Variant
    ID = pID
End Property
Property Let ID(value As Variant)
    pID = value
End Property

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("PermissionAddForm")
    AgregarForm.show
    Unload AgregarForm
    Set AgregarForm = Nothing
    If VarType(pID) = vbEmpty Then
        Me.getFields
        Call Form.fillListOrComboBox(permisos, Me.ListBox1)
    Else
        Me.getFieldsAndSelected
    End If
    
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub CommandButton3_Click()
    
    role.ID = pID
    role.name = Me.NameField
    permisosAsignados = Form.getSelectedItems(Me.ListBox1)
    
    If VarType(permisosAsignados) <> vbEmpty Then
        If VarType(pID) = vbEmpty Then
            role.create
            
            RoleHasPermission.roleID = role.ID
            
            For Each Item In permisosAsignados
                RoleHasPermission.permissionID = CInt(Item)
                Call RoleHasPermission.create
            Next Item
            
        Else
            RoleHasPermission.delete (pID)
            Call role.update(CInt(pID))
            
            RoleHasPermission.roleID = pID
            For Each Item In permisosAsignados
                RoleHasPermission.permissionID = CInt(Item)
                Call RoleHasPermission.create
            Next Item
            
        End If
    Else
        MsgBox "Seleccione al menos un permiso", vbInformation + vbOKOnly
        Exit Sub
    End If
    Me.Hide
    
End Sub

Private Sub UserForm_Activate()

    If VarType(pID) = vbEmpty Then
        Me.getFields
        Call Form.fillListOrComboBox(permisos, Me.ListBox1)
    Else
        Me.getFieldsAndSelected
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Set Permission = New cPermission
    Set RoleHasPermission = New cRoleHasPermission
    Set role = New cRole
    Set Form = New cForm
    
End Sub
Public Sub getFields()

    permisos = Permission.show( _
        fields:=Array("id", "name") _
    )
    
End Sub

Public Sub getFieldsAndSelected()
    
    Me.getFields
    Call Form.fillListOrComboBox(permisos, Me.ListBox1)
    
    joins = Array( _
        Array( _
            Array("role_has_permissions", "permissions"), _
            Array("permission_id", "id"), _
            "INNER" _
        ) _
    )
    permisosPorRol = RoleHasPermission.show( _
        fields:=Array("permissions.id", "permissions.name"), _
        colsFilter:=Array("role_id"), _
        logicOperators:=Array("="), _
        colsValues:=Array(pID), _
        joins:=joins _
    )

    Call Form.checkListBoxSelectedItems(permisosPorRol, permisos, Me.ListBox1)
    
End Sub




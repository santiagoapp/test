VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'
'                                   1. VARIABLES
'
'------------------------------------------------------------------------------------'

Private pStrCon As String
Private pDBPath As String
Private pQry As String
Private cn As ADODB.Connection
Private rs As ADODB.Recordset

'------------------------------------------------------------------------------------'
'
'                                   2. PROPIEDADES
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD DE LA CADENA DE CONEXION
'
'------------------------------------------------------------------------------------'

Public Property Get strCon() As String
    strCon = pStrCon
End Property

Public Property Let strCon(value As String)
    pStrCon = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.2 PROPIEDAD DE LA RUTA DE LA DB
'
'------------------------------------------------------------------------------------'

Public Property Get DBPath() As String
    DBPath = pDBPath
End Property

Public Property Let DBPath(value As String)
    pDBPath = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD DE LA QUERY
'
'------------------------------------------------------------------------------------'

Public Property Get qry() As String
    qry = pQry
End Property

Public Property Let qry(value As String)
    pQry = value
End Property

'------------------------------------------------------------------------------------'
'
'                                   3. CONSTRUCTOR
'
'------------------------------------------------------------------------------------'

Private Sub Class_Initialize()

End Sub

'------------------------------------------------------------------------------------'
'
'                                   4. MÉTODOS
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   4.1 REALIZA CONEXIÓN A LA BASE DE DATOS
'
'------------------------------------------------------------------------------------'

Public Function connect()

    Set cn = New ADODB.Connection
    pStrCon = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & pDBPath & ";"
    cn.Open pStrCon
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 AGREGA O ACTUALIZA UN REGISTRO DE LA BASE DE DATOS
'
'------------------------------------------------------------------------------------'

Public Function setRecords() As Boolean

On Error GoTo errorHandler
    Me.connect
    Set rs = New ADODB.Recordset
    Debug.Print pQry
    rs.Open pQry, cn, adOpenStatic
    setRecords = True
Exit Function
errorHandler:
    setRecords = False
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.3 OBTIENE REGISTROS DE LA BASE DE DATOS
'
'------------------------------------------------------------------------------------'

Public Function getRecords() As Variant
    
    Dim i As Integer
    Dim arr() As Variant
    
    Me.connect
    Set rs = New ADODB.Recordset
    Debug.Print pQry
    rs.Open pQry, cn, adOpenStatic
    
    If rs.EOF Then getRecords = Empty: Exit Function Else ReDim arr(rs.fields.count - 1, rs.RecordCount - 1)
    Do Until rs.EOF
        For j = 0 To rs.fields.count - 1
            arr(j, i) = rs.fields(j)
        Next
        i = i + 1
        rs.MoveNext
    Loop
    getRecords = arr
    
End Function

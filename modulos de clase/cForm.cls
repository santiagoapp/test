VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000

'Windows API calls to handle windows
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Public Sub RemoveCloseButton(objForm As Object)
    Dim lngStyle As LongPtr
    Dim lngHWnd As LongPtr

    Dim lpClassName As String
    lpClassName = vbNullString
    If val(Application.Version) >= 9 Then
       lpClassName = "ThunderDFrame"
    Else
       lpClassName = "ThunderXFrame"
    End If

    lngHWnd = FindWindow(lpClassName, objForm.Caption)
    lngStyle = GetWindowLongPtr(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLongPtr lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If
End Sub

Public Function fillListOrComboBox(ByVal arr As Variant, ByVal ListOrComboBox As Object)
    ListOrComboBox.Clear
    If VarType(arr) <> vbEmpty Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            ListOrComboBox.AddItem arr(0, j)
            For i = LBound(arr, 1) To UBound(arr, 1)
                If VarType(arr(i, j)) = vbNull Then
                    ListOrComboBox.List(j, i) = ""
                Else
                    ListOrComboBox.List(j, i) = arr(i, j)
                End If
            Next i
        Next j
    End If
End Function


Public Sub checkListBoxSelectedItems(ByVal selectedItems As Variant, ByVal items As Variant, ByVal ListOrComboBox As Object)
    
    If VarType(items) <> vbEmpty And VarType(selectedItems) <> vbEmpty Then
        ListOrComboBox.Clear
        For k = LBound(items, 2) To UBound(items, 2)
            ListOrComboBox.AddItem items(0, k)
            For i = LBound(items, 1) + 1 To UBound(items, 1)
                ListOrComboBox.List(k, i) = items(i, k)
            Next i
            For j = LBound(selectedItems, 2) To UBound(selectedItems, 2)
                If items(0, k) = selectedItems(0, j) Then
                    ListOrComboBox.Selected(k) = True
                    Exit For
                End If
            Next j
        Next k
    End If
End Sub


Public Function getSelectedItems(ByVal ListOrComboBox As Object) As Variant
    
    Dim arr As Variant
    Dim j As Integer
    ReDim arr(1000) As Variant
    j = 0
    
    For i = 0 To ListOrComboBox.ListCount - 1
        If ListOrComboBox.Selected(i) Then
            arr(j) = ListOrComboBox.List(i, 0)
            j = j + 1
        End If
    Next i
    If j <> 0 Then
        ReDim Preserve arr(j - 1)
        getSelectedItems = arr
        Exit Function
    End If
    ReDim Preserve arr(j)
    getSelectedItems = Empty
    
End Function


Public Function Calendar() As String
    
    Dim CalendarModal As Object
    Set CalendarModal = UserForms.Add("DatePicker")
    CalendarModal.show
    Calendar = CalendarModal.btnResponse
    Unload CalendarModal
    
End Function


Public Function fillMeses(ComboBox As Object)
    ComboBox.Clear
    For i = 1 To 12
        ComboBox.AddItem i
        ComboBox.List(i - 1, 1) = Format("01/" & Format(i, "00") & "/2000", "mmmm")
    Next i
    ComboBox.value = CInt(Month(Now))
End Function

Public Function fillAños(ComboBox As Object)
    ComboBox.Clear
    añoActal = year(Now)
    For i = añoActal - 10 To añoActal
        ComboBox.AddItem i
    Next i
    ComboBox.value = añoActal
End Function

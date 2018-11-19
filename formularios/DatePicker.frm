VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "Calendario"
   ClientHeight    =   3480
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   4320
   OleObjectBlob   =   "DatePicker.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Calendar1 As cCalendar
Attribute Calendar1.VB_VarHelpID = -1
Private valResponse As String
Private Form As cForm

Public Property Let btnResponse(btnResponse As String)
    valResponse = btnResponse
End Property
Public Property Get btnResponse() As String
    btnResponse = valResponse
End Property

Private Sub CommandButton2_Click()
    btnResponse = ""
    Me.Hide
End Sub

Private Sub CommandButton3_Click()
    btnResponse = Calendar1.value
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Set Form = New cForm
    Form.RemoveCloseButton Me
    Set Calendar1 = New cCalendar
    Calendar1.MonthLength = mlLocalShort
    Calendar1.BackColor = 16777215
    Calendar1.SelectedBackColor = 14671869
    Calendar1.HeaderBackColor = 3021805
    Calendar1.FirstDay = dwMonday
    Calendar1.GridFont = "Calibri"
    Calendar1.DayFont = "Calibri"
    Calendar1.TitleFont = "Calibri"
    Calendar1.TitleFontColor = 7762290
    Calendar1.SaturdayBackColor = 13092807
    Calendar1.SundayBackColor = 13092807
    Calendar1.Add_Calendar_into_Frame Me.Frame1
    
End Sub

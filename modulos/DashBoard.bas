Attribute VB_Name = "DashBoard"
Private Equipo As cEquipo
Private Correctivo As cCorrectivo
Private Form As cForm

Public Sub fillCards()
    
    Dim costoPreventivos As Double
    Dim costoCorrectivos As Double
    
    Set Equipo = New cEquipo
    Set Correctivo = New cCorrectivo
    
    totalEquipos = Equipo.count
    costoCorrectivos = Correctivo.sum("costo", Hoja8.ComboBox1, Hoja8.ComboBox2)
    Hoja8.Unprotect
        Hoja8.Shapes("card_1_value").TextFrame.Characters.Text = totalEquipos
        Hoja8.Shapes("card_4_value").TextFrame.Characters.Text = Format(costoCorrectivos, "$#,##0.00;-$#,##0.00")
        Hoja8.Shapes("card_5_value").TextFrame.Characters.Text = Format(costoCorrectivos + costoPreventivos, "$#,##0.00;-$#,##0.00")
    Hoja8.Protect
End Sub
Public Sub iniciar()

    Set Form = New cForm
    Call Form.fillMeses(Hoja8.ComboBox1)
    Call Form.fillAños(Hoja8.ComboBox2)
    Call Correctivos.getCorrectivosActivos
    Call DashBoard.fillCards
    
End Sub

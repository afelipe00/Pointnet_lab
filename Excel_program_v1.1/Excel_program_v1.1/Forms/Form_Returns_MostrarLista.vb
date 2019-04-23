Public Class Form_Returns_MostrarLista
    Public llenarlista As Form_return_purchase

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Dim Devolderdatos = New Class_add_return
        If Devolderdatos.ComprobarGarantia(ListView1) = "Valido" Then
            llenarlista.Label22.Text = Devolderdatos.Retornar_Valores(llenarlista.ComboBox2.Text, llenarlista.TextBox2, llenarlista.TextBox3, llenarlista.TextBox4,
                                       llenarlista.TextBox5, llenarlista.TextBox6, llenarlista.TextBox7, llenarlista.TextBox8,
                                       llenarlista.TextBox9, llenarlista.TextBox10, llenarlista.TextBox11, llenarlista.TextBox12)
            Me.Close()
        End If
        Label1.Text = "El tiempo de garantía del producto seleccionado ha vencido"
    End Sub

    Private Sub Form_Returns_MostrarLista_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Cargardatos = New Class_add_return
        Cargardatos.Llenarlista(llenarlista.ComboBox2.Text, ListView1)
        Label1.Text = "Para seleccionar el producto de doble clic sobre el"
    End Sub
End Class
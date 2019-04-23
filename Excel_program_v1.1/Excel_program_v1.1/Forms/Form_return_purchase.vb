Public Class Form_return_purchase
    Dim cantidad As String
    Private Sub Form_return_purchase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        Dim llenado = New Class_add_return
        TextBox1.Text = funcion.auto_codigo("Dev. Compra", "Mov. Inventario", 1)
        llenado.Comboboxllenar(ComboBox2, ComboBox3)
        Label8.Text = "Agregando devolución"
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim lista = New Form_Returns_MostrarLista
        Dim sumar = New Class_add_return
        lista.llenarlista = Me
        lista.ShowDialog()
        cantidad = Label22.Text
        Label22.Text = " "
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        Dim sumar = New Class_add_return
        If Label22.Text = " " Then
            sumar.cost_t(TextBox9, TextBox11, TextBox10, cantidad, Label8)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim escribir = New Class_add_return
        If escribir.ComprobarVacios(TextBox1, ComboBox2, TextBox2, TextBox5, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10, TextBox11, DateTimePicker1, ComboBox3, RichTextBox1, TextBox12, Label8) = 0 Then
            Me.Close()
        End If
    End Sub

    Private Sub ComboBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox2.KeyPress
        e.Handled = True
    End Sub

    Private Sub ComboBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox3.KeyPress
        e.Handled = True
    End Sub
End Class
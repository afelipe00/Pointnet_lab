Public Class Form_add_purchase
    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub Form_add_purchase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        TextBox1.Text = funcion.auto_codigo("Compra", "Mov. Inventario", 1)
        Label8.Text = "Agregando producto"
        funcion.combobox_lleno(ComboBox1, "Proveedores", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim llenar = New Class_add_compra
        llenar.llenar_textos(ComboBox1, TextBox2, TextBox3, TextBox4, "Proveedores")
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim agregar = New Form_add_item_compra
        Dim sumar = New Class_add_compra
        agregar.compra = Me
        agregar.ShowDialog()
        sumar.SumaCantidad(ListView1, TextBox5, 6)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim sumar = New Class_add_compra
        For Each ListViewItem In ListView1.SelectedItems
            ListViewItem.remove()
        Next
        sumar.SumaCantidad(ListView1, TextBox5, 6)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim agregar = New Class_add_compra
        If agregar.Comprobar_compra(TextBox1, DateTimePicker1, ComboBox1, TextBox2, TextBox3, TextBox4, ListView1, TextBox5, Label8, "Mov. Inventario") = 1 Then
            Close()
        End If
    End Sub

End Class
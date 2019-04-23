Public Class Form_sales
    Private Sub Form_sales_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        TextBox1.Text = funcion.auto_codigo("Venta", "Mov. Inventario", 1)
        Label8.Text = "Agregando productos"
        funcion.combobox_lleno(ComboBox1, "Clientes", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim llenar = New Class_add_venta
        llenar.llenar_textos(ComboBox1, TextBox2, TextBox3, TextBox4, "Clientes")
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim agregar = New Form_add_item_venta
        Dim sumar = New Class_add_venta
        agregar.venta = Me
        agregar.ShowDialog()
        sumar.SumaCantidad(ListView1, TextBox5, 6)
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim sumar = New Class_add_venta
        For Each ListViewItem In ListView1.SelectedItems
            ListViewItem.remove()
        Next
        sumar.SumaCantidad(ListView1, TextBox5, 6)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim agregar = New Class_add_venta
        If agregar.Comprobar_venta(TextBox1, DateTimePicker1, ComboBox1, TextBox2, TextBox3, TextBox4, ListView1, TextBox5, Label8, "Mov. Inventario") = 1 Then
            Close()
        End If
    End Sub
End Class
Public Class Form_add_item_venta
    Public venta As Form_sales
    Dim cantidadProducto As Integer
    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub Form_add_item_venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        Label12.Text = "Agregando producto"
        funcion.combobox_lleno(ComboBox1, "Inventario", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim llenar = New Class_add_venta
        cantidadProducto = llenar.llenar_textos_producto(ComboBox1, TextBox1, TextBox2, RichTextBox1, TextBox4, TextBox5, "Inventario")
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim costo = New Class_add_venta
        costo.cost_t(TextBox7, TextBox8, TextBox5, cantidadProducto, Label12)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim agregar = New Class_add_venta
        agregar.espacios_vacios(venta.ListView1, ComboBox1, TextBox7, TextBox1, TextBox2, RichTextBox1, TextBox4, TextBox5, TextBox8, DateTimePicker1, Label12)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
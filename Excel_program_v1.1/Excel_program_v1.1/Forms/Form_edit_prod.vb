Public Class Form_edit_prod

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para iniciar a modificar un producto
        Dim mod_prod = New Class_edit__prod
        mod_prod.caja_en_blanco(ComboBox1, TextBox2, TextBox3, RichTextBox1, TextBox5, TextBox6, TextBox7, Label8)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'boton para finalizar
        Dim finalizar = New Class_funciones
        Dim frm_finalizar = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_finalizar) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'boton para eliminar producto
        Dim mod_prod = New Class_edit__prod
        mod_prod.eliminar_prod(ComboBox1, TextBox2, TextBox3, RichTextBox1, TextBox5, TextBox6, TextBox7, "Productos", Label8)
    End Sub

    Private Sub Form_edit_prod_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, "Productos", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim edit_prod = New Class_edit__prod
        edit_prod.llenar_textos(ComboBox1, TextBox2, TextBox3, RichTextBox1, TextBox5, TextBox6, TextBox7, "Productos", Label8)
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub
End Class
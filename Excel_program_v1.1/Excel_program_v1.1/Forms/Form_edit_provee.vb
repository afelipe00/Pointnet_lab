Public Class Form_edit_provee

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'modificar el producto existente
        Dim edit_provee = New Class_edit_provee
        edit_provee.espacios_vacios(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, Label6, "Proveedores")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim finalizar = New Class_funciones
        Dim frm_finalizar = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_finalizar) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'eliminar producto 
        Dim eliminar = New Class_edit_provee
        eliminar.eliminar_prod(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, "Proveedores", Label6)
    End Sub

    Private Sub Form_edit_provee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'rellenar espacios del combobox
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, "Proveedores", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'rellenar espacios del formulario
        Dim rellenar = New Class_edit_provee
        rellenar.llenar_textos(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, "Proveedores", Label6)
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub
End Class
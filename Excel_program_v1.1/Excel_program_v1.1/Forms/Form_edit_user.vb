Public Class Form_edit_user
    Private Sub Form_edit_user_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, "Usuarios", 1, 2)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'boton para eliminar usuarios
        Dim elim_user = New Class_edit_user
        elim_user.eliminar_user(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, Label6, "Usuarios")
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        'boton para iniciar a modificar un producto
        Dim mod_user = New Class_edit_user
        mod_user.espacios_vacios(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, Label6, "Usuarios")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para finalizar
        Dim finalizar = New Class_edit_user
        Dim frm_finalizar = New Form_conf_finalizar
        Dim band As Boolean
        finalizar.conf_finalizar(frm_finalizar, band)
        If band = True Then
            Me.Close()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim edit_user = New Class_edit_user
        edit_user.llenar_textos(ComboBox1, TextBox2, TextBox3, TextBox4, TextBox5, "Usuarios", Label6)
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub
End Class
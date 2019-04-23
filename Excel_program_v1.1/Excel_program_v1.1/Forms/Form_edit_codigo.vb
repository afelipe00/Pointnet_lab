Public Class Form_edit_codigo
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'boton para finalizar registro
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim edit = New Class_edit_codigo
        edit.caja_en_blanco(ComboBox1, TextBox2, "Codigos", Label3)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim eliminar = New Class_edit_codigo
        eliminar.eliminar_prod(ComboBox1, TextBox2, "Codigos", Label3)
    End Sub

    Private Sub Form_edit_codigo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, "Codigos", 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim rellenar = New Class_edit_codigo
        rellenar.llenar_textos(ComboBox1, TextBox2, "Codigos", Label3)
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub
End Class
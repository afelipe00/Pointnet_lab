Public Class Form_add_user
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim agregar = New Class_add_user
        agregar.caja_en_blanco(TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, CheckBox1, CheckBox2, Label6, "Usuarios")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para finalizar
        Dim finalizar = New Class_funciones
        Dim frm_finalizar = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_finalizar) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_add_user_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label6.Text = "Agregando usuario"
    End Sub
End Class
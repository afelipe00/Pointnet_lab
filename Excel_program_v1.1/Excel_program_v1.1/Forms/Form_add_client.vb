Public Class Form_add_client

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim añadir = New Class_add_client
        añadir.caja_en_blanco(TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, Label6, "Clientes")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para finalizar registro
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_add_client_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label6.Text = "Registrando un nuevo cliente"
    End Sub
End Class
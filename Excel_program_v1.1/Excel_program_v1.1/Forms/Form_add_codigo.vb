Public Class Form_add_codigo
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim add_codigo = New Class_add_codigo
        add_codigo.caja_en_blanco(TextBox1, TextBox2, Label3, "Codigos")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para finalizar registro
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_add_codigo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = "Agregando nuevo codigo"
    End Sub
End Class
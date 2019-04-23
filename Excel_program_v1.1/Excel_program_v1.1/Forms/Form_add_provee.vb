Public Class Form_add_provee
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim add_prove = New Class_add_provee
        add_prove.caja_en_blanco(TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, Label6)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para finalizar registro
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_add_provee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label6.Text = "Agregando proveedor"
    End Sub
End Class
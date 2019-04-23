Public Class Form_add_prod

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'boton para registrar un producto
        Dim add_prod As New Class_add_prod
        add_prod.caja_en_blanco(TextBox1, TextBox2, TextBox3, RichTextBox1, TextBox5, TextBox6, TextBox7, Label8)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'boton para finalizar registro
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_add_prod_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, "Codigos", 1, 6)
        Label8.Text = "Registrando producto"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim funcion = New Class_funciones
        TextBox1.Text = funcion.auto_codigo(ComboBox1.Text, "Productos", 1)
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim frm_add_codigo = New Form_add_codigo
        frm_add_codigo.ShowDialog()
        ComboBox1.Items.Clear()
        funcion.combobox_lleno(ComboBox1, "Codigos", 1, 6)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim rib = Globals.Ribbons.Ribbon1
        Dim Workbook = Globals.ThisWorkbook.Application
        Dim frm_edit_codgio = New Form_edit_codigo
        If rib.Button4.Label = "Qbit" Then
            frm_edit_codgio.ShowDialog()
        End If
    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub
End Class
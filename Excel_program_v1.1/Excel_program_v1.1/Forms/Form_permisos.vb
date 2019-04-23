Public Class Form_permisos

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim modificar = New Class_user_rest
        modificar.espacios(ComboBox1, Label4)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim finalizar = New Class_funciones
        Dim frm_final = New Form_conf_finalizar
        If finalizar.confir_finalizar(frm_final) = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Form_permisos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim funcion = New Class_funciones
        funcion.combobox_lleno(ComboBox1, 8, 1, 2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim llenar = New Class_user_rest
        Dim boton() As CheckBox = {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8,
                                   CheckBox9, CheckBox10, CheckBox11, CheckBox12, CheckBox13, CheckBox14, CheckBox15,
                                   CheckBox16, CheckBox17, CheckBox18, CheckBox19, CheckBox20, CheckBox21}
        llenar.llenar_checkbox(ComboBox1, boton)
    End Sub
End Class
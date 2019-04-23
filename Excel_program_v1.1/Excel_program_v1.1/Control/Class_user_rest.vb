Public Class Class_user_rest
    Public Function espacios(ByRef user As ComboBox, ByRef label4 As Label)
        Dim funciones = New Class_funciones
        Dim fila As Long = funciones.repeat_fila(user.Text, "Usuarios", 2, 1, 1)
        If fila = 1 Then
            user.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label4.Text = "No existe este usuario"
        ElseIf user.Text = "" Then
            user.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label4.Text = "No pueden existir casillas vacias"
        Else
            mod_user(fila, label4)
            user.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        Return Nothing
    End Function
    Public Function mod_user(ByVal fila As Long, ByRef label4 As Label)
        Dim frmmod = New Form_permisos
        Dim boton() As Object = {frmmod.CheckBox1, frmmod.CheckBox2, frmmod.CheckBox3, frmmod.CheckBox4, frmmod.CheckBox5,
                                  frmmod.CheckBox6, frmmod.CheckBox7, frmmod.CheckBox8, frmmod.CheckBox9, frmmod.CheckBox10,
                                  frmmod.CheckBox11, frmmod.CheckBox12, frmmod.CheckBox13, frmmod.CheckBox14, frmmod.CheckBox15}
        Dim hojas() As Object = {frmmod.CheckBox16, frmmod.CheckBox17, frmmod.CheckBox18, frmmod.CheckBox19, frmmod.CheckBox20,
                                  frmmod.CheckBox21}
        Dim workbook = Globals.ThisWorkbook.Application
        For i = 0 To 14
            If boton(i).CheckState = 1 Then
                workbook.Sheets("Usuarios").cells(fila, i + 8).value = True
            Else
                workbook.Sheets("Usuarios").cells(fila, i + 8).value = False
            End If
        Next
        For i = 0 To 5
            If boton(i).CheckState = 1 Then
                workbook.Sheets("Usuarios").cells(fila, i + 23).value = True
            Else
                workbook.Sheets("Usuarios").cells(fila, i + 23).value = False
            End If
        Next
        If frmmod.CheckBox16.CheckState = 1 Then
            workbook.Sheets("Usuarios").cells(fila, 32).value = True
        Else
            workbook.Sheets("Usuarios").cells(fila, 32).value = True
        End If
        label4.Text = "Caracteristicas del usuario modificadas"
        Return Nothing
    End Function

    Public Function llenar_checkbox(ByRef user As ComboBox, ByRef boton() As CheckBox)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim frmmod = New Form_permisos
        Dim fila As Long = funciones.repeat_fila(user.Text, "Usuarios", 2, 1, 1)
        For i = 0 To boton.Length - 1
            If workbook.Sheets("Usuarios").cells(fila, 5).value = "Desarrollador" Then
            Else
                If workbook.Sheets("Usuarios").cells(fila, 8 + i).value = True Then
                    boton(i).Checked = True
                Else
                    boton(i).Checked = False
                End If
            End If
        Next
        Return Nothing
    End Function
End Class

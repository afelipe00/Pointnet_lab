Public Class Class_login
    Public Function ingresar(ByRef textbox1 As TextBox, ByRef textbox2 As TextBox, ByRef label1 As Label)
        Dim bulean = False
        Dim workbook = Globals.ThisWorkbook.Application
        Dim frm_ribbon = Globals.Ribbons.Ribbon1
        Dim funcion = New Class_funciones
        Dim exist = funcion.repeat_fila(textbox1.Text, "Usuarios", 2, 1, 1)
        Dim user As String = Convert.ToString(workbook.Sheets("Usuarios").cells(exist, 1).value)
        Dim contra As String = Convert.ToString(workbook.Sheets("Usuarios").cells(exist, 2).value)
        textbox2.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        textbox1.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        If exist <> 1 Then
            If user = textbox1.Text And contra = textbox2.Text Then
                workbook.Application.Visible = True
                frm_ribbon.Button4.Label = workbook.Sheets("Usuarios").cells(exist, 1).value
                bulean = True
                workbook.ActiveWorkbook.Protect("Qbit1234")
                workbook.Sheets(1).Protect("Qbit1234")
                workbook.Sheets(1).Select()
                workbook.ActiveWorkbook.Unprotect("Qbit1234")
                funcion.habilitar_user(frm_ribbon, user)
                workbook.ActiveWorkbook.Protect("Qbit1234")
            Else
                label1.Text = "Contraseña incorrecta"
                textbox2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            End If
        Else
            label1.Text = "Usuario incorrecto"
            textbox1.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
        End If
        Return bulean
    End Function

End Class

Public Class Form_Login

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim login = New Class_login
        Dim cerrar As Boolean
        cerrar = login.ingresar(UsernameTextBox, PasswordTextBox, Label1)
        If cerrar = True Then
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Dim workbook = Globals.ThisWorkbook
        workbook.Application.Quit()
        Me.Close()
    End Sub

    Private Sub Form_Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Ingresando usuario"
    End Sub

    Private Sub UsernameTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles UsernameTextBox.KeyPress
        If Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
End Class

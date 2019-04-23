Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        Dim login = New Form_Login
        Application.Visible = False
        login.ShowDialog()
    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

    End Sub

End Class

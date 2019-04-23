Public Class Form_conf_edi_prod

    Private press As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        press = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        press = False
        Me.Close()
    End Sub

    Public Property prop_press() As Boolean
        Get
            Return press
        End Get
        Set(value As Boolean)
            press = value
        End Set
    End Property

End Class
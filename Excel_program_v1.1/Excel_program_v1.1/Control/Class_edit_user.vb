Public Class Class_edit_user
    Public Function espacios_vacios(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox, ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label1 As Label,
                                    ByVal hoja As String)
        Dim vector() As Object = {combox, box2, box4, box3, box5}
        Dim estado As Byte = 0
        Dim funcion = New Class_funciones
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Or vector(i).text = "0" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label1.Text = "No pueden haber espacios en blanco"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        If funcion.Numcaracteres(box2.Text) < 8 Then
            box2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "No puede tener menos de 8 caracteres"
            estado = 1
        ElseIf box2.Text <> box3.Text Then
            box3.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            box2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Las contraseñas no coinciden"
            estado = 1
        End If

        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call mod_user(combox, box2, box3, box4, box5, label1, hoja)
        End If
        Return Nothing
    End Function
    Public Function mod_user(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                             ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label1 As Label, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        fila = funcion.repeat_fila(combox.Text, hoja, 1, 1, 1)
        'limpiar casilla en rojo 
        combox.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        'confiramcion para saber si los cuadros de precio de venta y compra no estan vacios
        'o tienen algun caracter que no debe tener precio
        If fila <> 1 Then
            With workbook.Sheets(hoja)
                .Cells(fila, 2).value = box2.Text
                .Cells(fila, 3).value = box4.Text
                .Cells(fila, 4).value = box5.Text
            End With
            limpiar_casillas(combox, box2, box3, box4, box5)
            label1.Text = "Producto modificado"
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If

        Return Nothing
    End Function

    Public Function conf_finalizar(ByRef final As Form_conf_finalizar, ByRef band As Boolean)
        final.ShowDialog()
        band = final.prop_press

        Return Nothing
    End Function

    Public Function llenar_textos(ByRef combox As ComboBox, ByRef text2 As TextBox, ByRef text3 As TextBox,
                                  ByRef text4 As TextBox, ByRef text5 As TextBox, ByVal hoja As String,
                                  ByRef label1 As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long
        'Buscar la fila repetida
        fila = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'Escribe en los cuadros de texto los valores de esa fila
        text2.Text = workbook.Sheets(hoja).cells(fila, 2).value
        text4.Text = workbook.Sheets(hoja).cells(fila, 3).value
        text5.Text = workbook.Sheets(hoja).cells(fila, 4).value
        label1.Text = "Modificando producto"
        Return Nothing
    End Function

    Public Function eliminar_user(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                             ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label1 As Label, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'limpiar casilla en rojo 
        combox.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        'verificar si el combobox tiene datos correctos
        If fila <> 1 Then
            'Elimina el producto
            workbook.Sheets(hoja).Rows(fila).delete()
            label1.Text = "Producto eliminado"
            'limpia el combo box de los productos eliminado
            combox.Items.Remove(combox.Text)
            combox.Refresh()
            'limpia todas las casillas del combo box
            limpiar_casillas(combox, box2, box3, box4, box5)
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function limpiar_casillas(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                                     ByRef box4 As TextBox, ByRef box5 As TextBox)
        combox.Text = ""
        box2.Text = ""
        box3.Text = ""
        box4.Text = ""
        box5.Text = ""
        combox.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        box3.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        box4.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        Return Nothing
    End Function
End Class

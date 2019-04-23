Public Class Class_edit_codigo
    Dim funcion = New Class_funciones

    Public Function caja_en_blanco(ByRef box1 As ComboBox, ByRef box2 As TextBox, ByVal hoja As String, ByRef label As Label)
        Dim vector() As Object = {box1, box2}
        Dim estado = 0
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                estado = 1
            ElseIf funcion.Numcaracteres(vector(1).text) > 4 Then
                vector(1).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No puede tener mas de 4 carácteres"
                estado = 1
            ElseIf funcion.Numcaracteres(vector(1).text) < 3 Then
                vector(1).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No puede tener menos de 3 carácteres"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call mod_producto(box1, box2, label, hoja)
        End If
        Return Nothing
    End Function
    Public Function mod_producto(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef label1 As Label, ByRef hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila As Long
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        fila = funcion.repeat_fila(combox.Text, hoja, 1, 1)
        'confiramcion para saber si los cuadros de precio de venta y compra no estan vacios
        'o tienen algun caracter que no debe tener precio
        If fila <> 1 Then
            With workbook.Sheets(hoja)
                .Cells(fila, 2).value = box2.Text
                .Cells(fila, 3).value = ribbon1.Button4.Label
            End With
            limpiar_casillas(combox, box2)
            label1.Text = "Producto modificado"
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function eliminar_prod(ByRef combox As ComboBox, ByRef box1 As TextBox, ByVal hoja As String, ByRef label1 As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'confirmar si el combobox no este vacio
        If fila <> 1 Then
            'Elimina el producto
            workbook.Sheets(hoja).Rows(fila).delete()
            label1.Text = "Producto eliminado"
            'limpia el combo box de los productos eliminados
            combox.Items.Remove(combox.Text)
            combox.Refresh()
            limpiar_casillas(combox, box1)
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function llenar_textos(ByRef combox As ComboBox, ByRef text1 As TextBox, ByVal hoja As String, ByRef label1 As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'Escribe en los cuadros de texto los valores de esa fila
        text1.Text = workbook.Sheets(hoja).cells(fila, 2).value
        label1.Text = "Modificando producto"
        Return Nothing
    End Function

    Public Function limpiar_casillas(ByRef combox As ComboBox, ByRef box1 As TextBox)
        combox.Text = ""
        box1.Text = ""
        Return Nothing
    End Function

    Public Function repeat_open(ByRef frm_conf_final As Form_conf_edi_prod)
        frm_conf_final.ShowDialog()
        Return frm_conf_final.prop_press
    End Function
End Class

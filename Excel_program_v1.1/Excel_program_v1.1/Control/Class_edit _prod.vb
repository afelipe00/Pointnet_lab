Public Class Class_edit__prod

    Public Function caja_en_blanco(ByRef box1 As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox, ByRef richbox1 As RichTextBox,
                                   ByRef box5 As TextBox, ByRef box6 As TextBox, ByRef box7 As TextBox, ByRef label As Label)
        Dim vector() As Object = {box1, box2, box3, box5, box6, box7}
        Dim funcion = New Class_funciones
        Dim estado = 0
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Or vector(i).text = "0" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                estado = 1
            ElseIf i > 3 And funcion.es_numero(vector(i).Text) = False Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "Los valores asignados no son numeros"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        'Condicion para verificar si son numeros o no son numeros los textos
        If funcion.es_numero(box6.Text) = True And funcion.es_numero(box7.Text) = True Then
            'condicion para que el precio de venta sea mayor al de comra
            If funcion.precio_coherente(box7.Text, box6.Text) = False Then
                box6.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                box7.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                estado = 1
            End If
        End If
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call mod_producto(box1, box2, box3, richbox1, box5, box6, box7, label, "Productos")
        End If
        Return Nothing
    End Function
    Public Function mod_producto(ByRef combox As ComboBox, ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As RichTextBox,
                                 ByRef box5 As TextBox, ByRef box6 As TextBox, ByRef box7 As TextBox, ByRef label1 As Label, ByRef hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim funcion = New Class_funciones
        Dim fila As Long
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        fila = funcion.repeat_fila(combox.Text, hoja, 1, 1, 1)
        'confiramcion para saber si los cuadros de precio de venta y compra no estan vacios
        'o tienen algun caracter que no debe tener precio
        If fila <> 1 Then
            With workbook.Sheets(hoja)
                .Cells(fila, 2).value = box1.Text
                .Cells(fila, 3).value = box2.Text
                .Cells(fila, 4).value = box3.Text
                .Cells(fila, 5).value = box5.Text
                .Cells(fila, 6).value = box6.Text
                .Cells(fila, 7).value = box7.Text
                .Cells(fila, 8).value = ribbon1.Button4.Label
            End With
            limpiar_casillas(combox, box1, box2, box3, box5, box6, box7)
            label1.Text = "Producto modificado"
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function eliminar_prod(ByRef combox As ComboBox, ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As RichTextBox,
                                  ByRef box5 As TextBox, ByRef box6 As TextBox, ByRef box7 As TextBox, ByVal hoja As String, ByRef label1 As Label)
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
            limpiar_casillas(combox, box1, box2, box3, box5, box6, box7)
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function llenar_textos(ByRef combox As ComboBox, ByRef text1 As TextBox, ByRef text2 As TextBox, ByRef text3 As RichTextBox,
                                  ByRef text4 As TextBox, ByRef text5 As TextBox, ByRef text6 As TextBox, ByVal hoja As String,
                                  ByRef label1 As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'Escribe en los cuadros de texto los valores de esa fila
        text1.Text = workbook.Sheets(hoja).cells(fila, 2).value
        text2.Text = workbook.Sheets(hoja).cells(fila, 3).value
        text3.Text = workbook.Sheets(hoja).cells(fila, 4).value
        text4.Text = workbook.Sheets(hoja).cells(fila, 5).value
        text5.Text = workbook.Sheets(hoja).cells(fila, 6).value
        text6.Text = workbook.Sheets(hoja).cells(fila, 7).value
        label1.Text = "Modificando producto"
        Return Nothing
    End Function

    Public Function limpiar_casillas(ByRef combox As ComboBox, ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As RichTextBox,
                                     ByRef box5 As TextBox, ByRef box6 As TextBox, ByRef box7 As TextBox)
        combox.Text = ""
        box1.Text = ""
        box2.Text = ""
        box3.Text = ""
        box5.Text = ""
        box6.Text = ""
        box7.Text = ""
        Return Nothing
    End Function

    Public Function repeat_open(ByRef frm_conf_final As Form_conf_edi_prod)
        frm_conf_final.ShowDialog()
        Return frm_conf_final.prop_press
    End Function
End Class

Public Class Class_edit_client
    Dim funcion = New Class_funciones
    Public Function espacios_vacios(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                                    ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label1 As Label,
                                    ByVal hoja As String)
        Dim vector() As Object = {combox, box2, box4, box3, box5}
        Dim estado As Byte = 0
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label1.Text = "No pueden haber espacios en blanco"
                estado = 1
            ElseIf i > 2 And funcion.es_numero(vector(i).Text) = False Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label1.Text = "Los valores asignados no son numeros"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call mod_client(combox, box2, box3, box4, box5, label1, hoja)
        End If
        Return Nothing
    End Function
    Public Function mod_client(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                               ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label1 As Label,
                               ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila, rep As Long
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        fila = funcion.repeat_fila(combox.Text, hoja, 1, 1, 1)
        rep = funcion.repetir_valor(fila, hoja, box3.Text, box2.Text, box3.Text, box4.Text, 3, 2, 3, 4)
        'limpiar casilla en rojo 
        combox.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        'confiramcion para saber si los cuadros de precio de venta y compra no estan vacios
        'o tienen algun caracter que no debe tener precio
        If fila <> 1 Then
            If fila = rep Then
                With workbook.Sheets(hoja)
                    .Cells(fila, 2).value = box2.Text
                    .Cells(fila, 3).value = box3.Text
                    .Cells(fila, 4).value = box4.Text
                    .Cells(fila, 5).value = box5.Text
                    .Cells(fila, 6).value = ribbon1.Button4.Label
                End With
                limpiar_casillas(combox, box2, box3, box4, box5)
                label1.Text = "Producto modificado"
            Else
                label1.Text = "Estas especificaciones existen en otro producto"
            End If
        Else
            combox.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label1.Text = "Ingrese un codigo de producto existente"
        End If
        Return Nothing
    End Function

    Public Function eliminar_client(ByRef combox As ComboBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                                    ByRef box4 As TextBox, ByRef box5 As TextBox, ByVal hoja As String,
                                    ByRef label1 As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'limpiar casilla en rojo 
        combox.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        'verificar si el combobox tiene datos correctos
        If combox.Text <> "" And fila <> 1 Then
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
        text3.Text = workbook.Sheets(hoja).cells(fila, 3).value
        text4.Text = workbook.Sheets(hoja).cells(fila, 4).value
        text5.Text = workbook.Sheets(hoja).cells(fila, 5).value
        label1.Text = "Modificando producto"
        Return Nothing
    End Function

    Public Function limpiar_casillas(ByRef combox As ComboBox, ByRef box2 As TextBox,
                                     ByRef box3 As TextBox, ByRef box4 As TextBox, ByRef box5 As TextBox)
        combox.Text = ""
        box2.Text = ""
        box3.Text = ""
        box4.Text = ""
        box5.Text = "0"
        Return Nothing
    End Function

    Public Function repeat_open(ByRef frm_conf_final As Form_conf_edi_prod)
        frm_conf_final.ShowDialog()
        Return frm_conf_final.prop_press
    End Function
End Class
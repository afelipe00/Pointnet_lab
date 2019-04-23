Public Class Class_add_venta
    Public Function espacios_vacios(ByRef lista As ListView, ByRef combox As ComboBox, ByRef cantidad As TextBox, ByRef nom As TextBox,
                                    ByRef ref As TextBox, ByRef Descripcion As RichTextBox, ByRef Fabricante As TextBox, ByRef costu As TextBox,
                                    ByRef costt As TextBox, ByRef garantia As DateTimePicker, ByRef label As Label)
        Dim vector() As Object = {combox, cantidad, costt}
        Dim funcion = New Class_funciones
        Dim contador As Byte = 0
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Or vector(i).text = "0" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                contador = 1
            ElseIf i > 1 And funcion.es_numero(vector(i).Text) = False Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "Los valores asignados no son numeros"
                contador = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        If contador = 0 Then
            Call Agregar_lista(lista, cantidad, combox, nom, ref, costu, costt, garantia)
            Call LimpiarCasillas(combox, cantidad, nom, ref, costu, costt, Fabricante, Descripcion)
        End If
        Return contador
    End Function

    Public Function Comprobar_venta(ByRef codigo As TextBox, ByRef fecha As DateTimePicker, ByRef id As ComboBox, ByRef cliente As TextBox,
                                     ByRef descuento As TextBox, ByRef correo As TextBox, ByRef lista As ListView, ByRef total As TextBox,
                                     ByRef label1 As Label, ByVal hoja As String)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim vector() As Object = {cliente, id, descuento, correo, lista, total}
        Dim estado As Byte = 0
        For i = 0 To vector.Length - 1
            If vector(i).text = "" And i < 4 Then
                estado = 1
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            ElseIf i > 3 And lista.Items.Count = 0 Then
                estado = 1
                vector(vector.Length - 1).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        If estado = 0 Then
            Agregar_venta(codigo, fecha, cliente.Text, lista, hoja)
            Dim fil_celular As Long = funciones.repeat_fila(id.Text, "Clientes", 2, 1, 1)
            Dim celular As Long = workbook.Sheets("Clientes").cells(fil_celular, 3).value
            funciones.Llenar_Factura("Form. Facturas", codigo.Text, fecha.Text, cliente.Text, id.Text, correo.Text, celular, descuento.Text, lista)
            funciones.Guardar_Factura("Form. Facturas", codigo.Text, fecha.Text, lista)
            Agregar_inventarios(lista, "Inventario", 1, 6)
            Return 1
        Else
            label1.Text = "No pueden haber espacios vacios"
            Return Nothing
        End If
        funciones.Limpiar_Factura(lista)
    End Function

    Public Function Agregar_venta(ByRef codigo As TextBox, ByRef fecha As DateTimePicker, ByRef cliente As String, ByRef lista As ListView,
                                   ByRef hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon = Globals.Ribbons.Ribbon1
        Dim funcion = New Class_funciones
        Dim consecutivo() As String
        Dim fila As Long
        consecutivo = codigo.Text.Split("-")
        fila = funcion.recorrer_filas(1, hoja, 1)
        'verificar que esten repetidos
        For i = 0 To lista.Items.Count - 1
            With workbook.Sheets(hoja)
                .Cells(fila + i, 1).value = codigo.Text
                .Cells(fila + i, 2).value = fecha.Text
                .Cells(fila + i, 3).value = "VENTA"
                .Cells(fila + i, 4).value = lista.Items(i).SubItems(1).Text & " " & lista.Items(i).SubItems(2).Text
                .Cells(fila + i, 5).value = consecutivo(1)
                .Cells(fila + i, 6).value = cliente
                .Cells(fila + i, 7).value = lista.Items(i).SubItems(3).Text
                .Cells(fila + i, 8).value = lista.Items(i).Text
                .Cells(fila + i, 9).value = "N/A"
                .Cells(fila + i, 10).value = lista.Items(i).SubItems(4).Text
                .Cells(fila + i, 11).value = lista.Items(i).SubItems(5).Text
                .Cells(fila + i, 12).value = ribbon.Button4.Label
            End With
        Next
        Return Nothing
    End Function

    Public Function Agregar_lista(ByRef lista As ListView, ByRef cantidad As TextBox, ByRef codigo As ComboBox, ByRef nom As TextBox,
                                  ByRef ref As TextBox, ByRef costu As TextBox, ByRef costt As TextBox, ByRef garantia As DateTimePicker)
        With lista.Items.Add(cantidad.Text)
            .SubItems.Add(codigo.Text)
            .SubItems.Add(nom.Text & "-" & ref.Text)
            .SubItems.Add(garantia.Text)
            .SubItems.Add(costu.Text)
            .SubItems.Add(costt.Text)
        End With
        Return Nothing
    End Function

    Public Function cost_t(ByRef cantidad As TextBox, ByRef costo_total As TextBox, ByRef costo_u As TextBox, ByVal cantidadmax As Integer,
                           ByRef label As Label)
        Dim funcion = New Class_funciones
        If funcion.es_numero(cantidad.Text) = True Then
            If cantidad.Text <= cantidadmax Then
                costo_total.Text = cantidad.Text * costo_u.Text
            Else
                costo_total.Text = "0"
            End If
        Else
            costo_total.Text = "0"
        End If
        Return Nothing
    End Function

    Public Function llenar_textos(ByRef combox As ComboBox, ByRef text2 As TextBox, ByRef text3 As TextBox, ByRef text4 As TextBox,
                                  ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila As Long
        'Buscar la fila repetida
        fila = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        'Escribe en los cuadros de texto los valores de esa fila
        text2.Text = workbook.Sheets(hoja).cells(fila, 2).value
        text3.Text = workbook.Sheets(hoja).cells(fila, 5).value
        text4.Text = workbook.Sheets(hoja).cells(fila, 4).value
        Return Nothing
    End Function

    Public Function llenar_textos_producto(ByRef combox As ComboBox, ByRef text1 As TextBox, ByRef text2 As TextBox, ByRef text3 As RichTextBox,
                                           ByRef text4 As TextBox, ByRef text5 As TextBox, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim cantidad As Integer
        Dim fila As Long
        'Buscar la fila repetida
        fila = funcion.repeat_fila(combox.Text, hoja, 2, 1, 1)
        cantidad = workbook.Sheets(hoja).cells(fila, 6).value
        'Escribe en los cuadros de texto los valores de esa fila
        text1.Text = workbook.Sheets(hoja).cells(fila, 2).value
        text2.Text = workbook.Sheets(hoja).cells(fila, 3).value
        text3.Text = workbook.Sheets(hoja).cells(fila, 4).value
        text4.Text = workbook.Sheets(hoja).cells(fila, 5).value
        text5.Text = workbook.Sheets(hoja).cells(fila, 8).value
        Return cantidad
    End Function

    Public Function Agregar_inventarios(ByRef lista As ListView, ByVal hoja As String, ByVal colum_codigo As Integer, ByVal colum_cantidad As Integer)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim codigos(lista.Items.Count) As String
        Dim cantidad(lista.Items.Count) As String
        Dim fila_final, fila_rep As Long
        For i = 0 To lista.Items.Count - 1
            'llenando vectores
            codigos(i) = lista.Items(i).SubItems(1).Text
            cantidad(i) = lista.Items(i).Text
            'buscando esos valores en la hoja de invenatior
            fila_final = funciones.recorrer_filas(1, hoja, 1)
            fila_rep = funciones.repeat_fila(codigos(i), hoja, 2, colum_codigo, 1)
            If fila_rep <> 1 Then
                workbook.Sheets(hoja).Cells(fila_rep, colum_cantidad).value = workbook.Sheets(hoja).Cells(fila_rep, colum_cantidad).value - cantidad(i)
                workbook.Sheets(hoja).Cells(fila_rep, colum_cantidad + 3).value = workbook.Sheets(hoja).Cells(fila_rep, colum_cantidad).value * workbook.Sheets(hoja).Cells(fila_rep, colum_cantidad + 1).value
            End If
        Next
        Return Nothing
    End Function

    Public Function SumaCantidad(ByRef ListView1 As ListView, ByRef textbox As TextBox, ByVal columna As Integer)
        Dim suma As Integer = 0
        For i = 0 To ListView1.Items.Count - 1
            suma = suma + CInt(ListView1.Items(i).SubItems(columna - 1).Text)
        Next
        textbox.Text = suma
        Return Nothing
    End Function

    Public Function LimpiarCasillas(ByRef ComboBox1 As ComboBox, ByRef TextBox1 As TextBox, ByRef TextBox2 As TextBox, ByRef TextBox4 As TextBox,
                                    ByRef TextBox5 As TextBox, ByRef TextBox7 As TextBox, ByRef TextBox8 As TextBox, ByRef RichTextBox1 As RichTextBox)
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        RichTextBox1.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        Return Nothing
    End Function
End Class

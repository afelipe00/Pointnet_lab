Public Class Class_add_return
    Public Function ComprobarGarantia(ByRef lista As ListView)
        Dim estado As String=""
        For Each ListViewItem In lista.SelectedItems
            If DateTime.Now.ToString("dd/MM/yyyy") <= ListViewItem.SubItems(3).text Then
                estado = "Valido"
            Else
                estado = Nothing
            End If
        Next
        Return estado
    End Function

    Public Function ComprobarVacios(ByRef codigodevolucion As TextBox, ByRef codigofactura As ComboBox, ByRef provee As TextBox, ByRef telefono As TextBox,
                                    ByRef email As TextBox, ByRef codigo_producto As TextBox, ByRef nombre_producto As TextBox, ByRef cantidad As TextBox,
                                    ByRef precio_unitario As TextBox, ByRef precio_total As TextBox, ByRef fecha As DateTimePicker, ByRef estado As ComboBox,
                                    ByRef causal_dev As RichTextBox, ByRef responsablecompra As TextBox, ByRef label As Label)
        Dim vector() As Object = {codigofactura, estado, causal_dev}
        Dim funcion = New Class_funciones
        Dim contador As Byte = 0
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                contador = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        If contador = 0 Then
            Call Retornar(codigodevolucion.Text, codigofactura.Text, provee.Text, telefono.Text, email.Text, codigo_producto.Text, nombre_producto.Text,
                          cantidad.Text, precio_unitario.Text, precio_total.Text, fecha.Text, estado.Text, causal_dev.Text, responsablecompra.Text)
        End If
        Return contador
    End Function

    Public Function Comboboxllenar(ByRef combox As ComboBox, ByRef combox2 As ComboBox)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim fila_final = funciones.recorrer_filas(1, "Mov. Inventario", 1)
        Dim repetida As String = ""
        For i = 1 To fila_final
            If workbook.Sheets("Mov. Inventario").cells(i, 3).value = "COMPRA" Then
                If repetida = Convert.ToString(workbook.Sheets("Mov. Inventario").cells(i, 1).value) Then

                Else
                    combox.Items.Add(Convert.ToString(workbook.Sheets("Mov. Inventario").cells(i, 1).value))
                    repetida = Convert.ToString(workbook.Sheets("Mov. Inventario").cells(i, 1).value)
                End If
            End If
        Next
        combox2.Items.Add("EN PROCESO")
        combox2.Items.Add("FINALIZADO")
        combox2.Items.Add("PERDIDA")
        Return Nothing
    End Function

    Public Function Llenarlista(ByRef combox As String, ByRef Lista As ListView)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim fila_final = funciones.recorrer_filas(1, "Mov. Inventario", 1)
        For i = 1 To fila_final
            If workbook.Sheets("Mov. Inventario").cells(i, 1).value = combox Then
                With Lista.Items.Add(workbook.Sheets("Mov. Inventario").cells(i, 8).value)
                    .SubItems.Add(workbook.Sheets("Mov. Inventario").cells(i, 1).value)
                    .SubItems.Add(workbook.Sheets("Mov. Inventario").cells(i, 4).value)
                    .SubItems.Add(workbook.Sheets("Mov. Inventario").cells(i, 7).value)
                    .SubItems.Add(workbook.Sheets("Mov. Inventario").cells(i, 10).value)
                    .SubItems.Add(workbook.Sheets("Mov. Inventario").cells(i, 11).value)
                End With
            End If
        Next
        Return Nothing
    End Function

    Public Function Retornar_Valores(ByVal codigo As String, ByRef provee As TextBox, ByRef repre As TextBox, ByRef Nit As TextBox, ByRef telefono As TextBox,
                                    ByRef email As TextBox, ByRef codigo_producto As TextBox, ByRef nombre_producto As TextBox, ByRef cantidad As TextBox,
                                    ByRef costounitario As TextBox, ByRef total As TextBox, ByRef responsable As TextBox)
        Dim funciones = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim fila = funciones.repeat_fila(codigo, "Mov. Inventario", 2, 1, 1)
        Dim otherfila = fila
        Dim tercero, producto As String
        Dim separador() As String
        If fila <> 1 Then
            tercero = workbook.Sheets("Mov. Inventario").cells(fila, 6).value
            producto = workbook.Sheets("Mov. Inventario").cells(fila, 4).value
            fila = funciones.repeat_fila(tercero, "Proveedores", 2, 1, 1) 'La fila en la que se encuentra el proveedor de la devolución
            If fila <> 1 Then
                provee.Text = Convert.ToString(workbook.Sheets("Proveedores").cells(fila, 1).value)
                repre.Text = Convert.ToString(workbook.Sheets("Proveedores").cells(fila, 2).value)
                Nit.Text = Convert.ToString(workbook.Sheets("Proveedores").cells(fila, 4).value)
                telefono.Text = Convert.ToString(workbook.Sheets("Proveedores").cells(fila, 3).value)
                email.Text = Convert.ToString(workbook.Sheets("Proveedores").cells(fila, 5).value)
                responsable.Text = workbook.Sheets("Mov. Inventario").cells(fila, 12).value
            End If
            separador = producto.Split(" ") 'Esto es necesario para separar un texto en el que se encuentra tanto el codigo del producto como otras cosas
            fila = funciones.repeat_fila(separador(0), "Productos", 2, 1, 1) 'La fila en la que se encuentra el producto de la devolución
            If fila <> 1 Then
                codigo_producto.Text = Convert.ToString(workbook.Sheets("Productos").cells(fila, 1).value)
                nombre_producto.Text = Convert.ToString(workbook.Sheets("Productos").cells(fila, 2).value)
                costounitario.Text = Convert.ToString(workbook.Sheets("Productos").cells(fila, 6).value)
                total.Text = Convert.ToString(workbook.Sheets("Mov. Inventario").cells(otherfila, 11).value)
                cantidad.Text = Convert.ToString(workbook.Sheets("Mov. Inventario").cells(otherfila, 8).value)
                responsable.Text = workbook.Sheets("Mov. Inventario").cells(fila, 12).value
            End If
        End If
        Return cantidad.Text
    End Function

    Public Function cost_t(ByRef cantidad As TextBox, ByRef costo_total As TextBox, ByRef costo_u As TextBox, ByRef limite_cantidad As String,
                           ByRef label As Label)
        Dim funcion = New Class_funciones
        If funcion.es_numero(cantidad.Text) = True Then
            If cantidad.Text <= limite_cantidad Then
                costo_total.Text = cantidad.Text * costo_u.Text
                label.Text = "Agregando devolución"
            Else
                label.Text = "La cantidad de productos que desea devolver es superior a las compradas"
                costo_total.Text = ""
            End If
        Else
            costo_total.Text = ""
            label.Text = "Debera escoger una cantidad de productos valida que desea devolver"
        End If
        Return Nothing
    End Function

    Public Function Retornar(ByVal codigodevolucion As String, ByVal codigofactura As String, ByVal proveeoid As String, ByVal telefono As String,
                             ByVal email As String, ByVal codigo_producto As String, ByVal nombre_producto As String, ByVal cantidad As String,
                             ByVal precio_unitario As String, ByVal precio_total As String, ByVal fecha As String, ByVal estado As String,
                             ByVal causal_dev As String, ByVal responsablecompra As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon = Globals.Ribbons.Ribbon1
        Dim funciones = New Class_funciones
        Dim fila_final = funciones.recorrer_filas(1, "Devoluciones", 1)
        workbook.Sheets("Devoluciones").Cells(fila_final, 1).value = codigodevolucion
        workbook.Sheets("Devoluciones").Cells(fila_final, 2).value = fecha
        workbook.Sheets("Devoluciones").Cells(fila_final, 3).value = estado
        workbook.Sheets("Devoluciones").Cells(fila_final, 4).value = codigofactura
        workbook.Sheets("Devoluciones").Cells(fila_final, 5).value = codigo_producto & " " & nombre_producto
        workbook.Sheets("Devoluciones").Cells(fila_final, 6).value = cantidad
        workbook.Sheets("Devoluciones").Cells(fila_final, 7).value = precio_unitario
        workbook.Sheets("Devoluciones").Cells(fila_final, 8).value = precio_total
        workbook.Sheets("Devoluciones").Cells(fila_final, 9).value = causal_dev
        workbook.Sheets("Devoluciones").Cells(fila_final, 10).value = responsablecompra
        workbook.Sheets("Devoluciones").Cells(fila_final, 11).value = proveeoid
        workbook.Sheets("Devoluciones").Cells(fila_final, 12).value = telefono
        workbook.Sheets("Devoluciones").Cells(fila_final, 13).value = email
        workbook.Sheets("Devoluciones").Cells(fila_final, 14).value = ribbon.Button4.Label
        Estado_devolucion(estado, codigodevolucion, codigo_producto, cantidad, fecha, proveeoid, "N/A", email, telefono, "0")
        Return Nothing
    End Function
    'TODAS ESTAS FUNCIONES HACEN FALTA TERMINARLAS
    Public Function Estado_devolucion(ByVal estado As String, ByRef codigodevolucion As String, ByVal codigoproducto As String, ByVal Cantidadproducto As Integer,
                                      ByVal fecha As String, ByVal tercero As String, ByVal id As String, ByVal correo As String, ByVal celular As String, ByVal Descuento As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funciones = New Class_funciones
        Dim filaproducto As Long = funciones.repeat_fila(codigoproducto, "Inventario", 1, 1, 1)
        If estado = "FINALIZADO" Then
            Call Llenar_Factura("Form. Facturas", codigodevolucion, fecha, tercero, id, correo, celular, Descuento)
            Guardar_Factura("Form. Facturas", codigodevolucion, fecha)
        ElseIf estado = "PERDIDA" Then
            Call Llenar_Factura("Form. Facturas", codigodevolucion, fecha, tercero, id, correo, celular, Descuento)
            Call Guardar_Factura("Form. Facturas", codigodevolucion, fecha)

            workbook.Sheets("Inventario").Cells(filaproducto, 6).value = workbook.Sheets("Inventario").Cells(filaproducto, 6).value - Cantidadproducto
            workbook.Sheets("Inventario").Cells(filaproducto, 9).value = workbook.Sheets("Inventario").Cells(filaproducto, 6).value * workbook.Sheets("Inventario").Cells(filaproducto, 7).value
        End If
        Return Nothing
    End Function

    Public Function Llenar_Factura(ByVal hoja As String, ByVal codigo As String, ByVal Fecha As String, ByVal nombre As String, ByVal id As String,
                                   ByVal email As String, ByVal celular As Long, ByVal Descuento As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim tot As Long = 0
        Dim f_list As Long = funcion.repeat_fila(codigo, "Devoluciones", 1, 1, 1)
        Dim iva As Integer = 19
        Dim consecutivo() As String = codigo.Split("-")
        Dim nom_ref() As String = workbook.Sheets("Devoluciones").Cells(f_list, 5).value.split(" ")
        Dim filaproducto As Long = funcion.repeat_fila(nom_ref(0), "Productos", 1, 1, 1)
        With workbook.Sheets(hoja)
            .Cells(1, 6).value = consecutivo(1)
            .Cells(3, 6).value = Fecha
            .Cells(10, 2).value = nombre
            .Cells(11, 5).value = id
            .Cells(11, 2).value = email
            .Cells(12, 5).value = celular
            .Cells(12, 2).value = "Bogotá D.C."
            If f_list <> 1 Then
                .Cells(15, 1).value = "1"
                .Cells(15, 2) = workbook.Sheets("Devoluciones").Cells(f_list, 6).value
                .cells(15, 3) = workbook.Sheets("Productos").Cells(filaproducto, 2).value
                .cells(15, 4) = workbook.Sheets("Productos").Cells(filaproducto, 3).value
                .cells(15, 5) = workbook.Sheets("Productos").Cells(filaproducto, 7).value
                .cells(15, 6) = .cells(15, 2).value * .cells(15, 5).value
            End If
            .Cells(18, 1).Value = "La garantia de devolución consta de la misma que se ha impuesto anteriormente"
            .Cells(19, 6).value = Descuento
            .Cells(20, 6).value = iva
            .Cells(23, 2).value = "Los módulos electrónicos tiene garantía de 10 días después de su compra, y los componentes electrónicos solamente 5 días, todo ello por defectos de fabrica"
            .Cells(26, 3).value = "Contado-efectivo"
            .Cells(27, 3).value = "Inmediato"
            tot = tot + .cells(15, 6).value
            .cells(15 + 3, 6).value = tot
            tot = tot * (iva / 100) + tot
            .cells(15 + 6, 6).value = tot - (tot * (Descuento / 100))
        End With
        Return Nothing
    End Function
    Public Function Limpiar_Factura()
        Dim workbook = Globals.ThisWorkbook.Application
        For i = 0 To 5
            Workbook.Sheets("Form. Facturas").cells(15, 1 + i).value = ""
        Next
        With Workbook.Sheets("Form. Facturas")
            .Cells(1, 6).value = ""
            .Cells(3, 6).value = ""
            .Cells(10, 2).value = ""
            .Cells(11, 5).value = ""
            .Cells(11, 2).value = ""
            .Cells(12, 5).value = ""
            .Cells(12, 2).value = ""
            .cells(18, 6).value = ""
            .Cells(19, 6).value = ""
            .Cells(20, 6).value = ""
            .Cells(21, 6).value = ""
            .Cells(22, 6).value = ""
            .Cells(23, 2).value = ""
            .Cells(26, 3).value = ""
            .Cells(27, 3).value = ""
        End With
        Return Nothing
    End Function

    Public Function Guardar_Factura(ByVal hoja As String, ByRef codigo As String, ByRef Fecha As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ubicacion As String
        Dim Compraoventa() As String
        Fecha = Replace(Fecha, "/", "-")
        Compraoventa = codigo.Split("-")
        If Compraoventa(0) = "VTA" Then
            ubicacion = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\QbitElectronica\Facturas Venta\"
        ElseIf Compraoventa(0) = "CMP" Then
            ubicacion = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\QbitElectronica\Facturas Compra\"
        Else
            ubicacion = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\QbitElectronica\Facturas Devolución\"
        End If
        My.Computer.FileSystem.CreateDirectory(ubicacion)
        workbook.Sheets(hoja).PrintOut(PrintToFile:=True, PrToFileName:=ubicacion & codigo & " " & Fecha & ".pdf", ActivePrinter:="Microsoft Print to PDF")
        ubicacion = ubicacion & codigo & " " & Fecha & ".pdf"
        Call Limpiar_Factura()
        'MOSTRAR FACTURA AUN NO ARREGLADO
        'If Not My.Computer.FileSystem.DirectoryExists(ubicacion) Then
        'Diagnostics.Process.Start(ubicacion)
        'End If
        Return Nothing
    End Function
End Class

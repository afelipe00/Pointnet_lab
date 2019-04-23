Public Class Class_funciones
    Dim workbook = Globals.ThisWorkbook.Application
    Public Function recorrer_filas(ByRef fila As Long, ByRef hoja As String, ByVal colum As Integer) As Long
        Do While Convert.ToString(workbook.Sheets(hoja).Cells(fila, colum).value) <> ""
            fila = fila + 1
        Loop
        Return fila
    End Function

    Public Function repetir_valor(ByVal fila As Long, ByRef hoja As String, ByRef text1 As String,
                                  ByRef text2 As String, ByRef text3 As String, ByRef text4 As String,
                                  ByVal colum As Integer, ByVal colum1 As Integer, ByVal colum2 As Integer,
                                  ByVal colum3 As Integer)
        Dim i As Long
        For i = 1 To fila
            If Convert.ToString(workbook.sheets(hoja).cells(i, colum).value) = text1 Then
                Exit For
            End If
            If Convert.ToString(workbook.Sheets(hoja).cells(i, colum1).value) = text2 Then
                If Convert.ToString(workbook.Sheets(hoja).cells(i, colum2).value) = text3 Then
                    If Convert.ToString(workbook.Sheets(hoja).cells(i, colum3).value) = text4 Then
                        Exit For
                    End If
                End If
            End If
        Next
        Return i - 1
    End Function

    Public Function es_numero(ByVal texto As String) As Boolean
        Dim arraytext() = texto.ToCharArray()
        Dim band = True
        If texto <> "" Then
            For i = 0 To arraytext.Length - 1
                If Char.IsNumber(arraytext, i) = False And arraytext(i) <> "." And arraytext(i) <> "," Then
                    band = False
                    Exit For
                End If
            Next
        Else
            band = False
        End If
        Return band
    End Function

    Public Function combobox_lleno(ByRef combox As ComboBox, ByVal hoja As String, ByVal colum As Integer, ByVal fila_init As Long)
        Dim fila As Long
        'inicializar rango para agregar el producto
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        fila = recorrer_filas(fila, hoja, colum)
        For i = fila_init To fila - 1
            If Convert.ToString(workbook.sheets(hoja).cells(i, 5).value) = "Desarrollador" Then
            Else
                combox.Items.Add(Convert.ToString(workbook.Sheets(hoja).cells(i, colum).value))
            End If
        Next
        Return Nothing
    End Function

    Public Function repeat_fila(ByVal combox As String, ByVal hoja As String, ByVal fila_init As Long, ByVal columna As Integer, ByVal opcion As Integer)
        Dim fila As Long
        Dim bandera As Integer = 0
        Dim final As Long
        fila = recorrer_filas(1, hoja, columna)
        For final = fila_init To fila - 1
            If Convert.ToString(workbook.Sheets(hoja).cells(final, columna).value) = combox Then
                bandera = 1
                Exit For
            End If
        Next
        If opcion = 1 Then
            If bandera = 1 Then
                Return final
            Else
                Return 1
            End If
        Else
            Return final
        End If
    End Function

    Public Function precio_coherente(ByVal valor_ven As Integer, ByVal valor_com As Integer)
        Dim bulean = True
        If valor_ven < valor_com Then
            bulean = False
        End If
        Return bulean
    End Function

    Public Function confir_finalizar(ByRef frm_finalizar As Form_conf_finalizar)
        'inicializacion de form para confirmar la finalizacion de el add producto
        frm_finalizar.ShowDialog()
        Return frm_finalizar.prop_press
    End Function

    Public Function habilitar_user(ByRef ribbon As Ribbon1, ByVal user As String)
        Dim botones() As Object = {ribbon.Button1, ribbon.Menu4, ribbon.Menu1, ribbon.Menu2, ribbon.Menu5, ribbon.Menu3,
                                   ribbon.Button5, ribbon.Button6, ribbon.Button16, ribbon.Button17, ribbon.Button15,
                                   ribbon.Button22, ribbon.Button21, ribbon.Button20, ribbon.Button23, ribbon.Button24,
                                   ribbon.Button25}
        Dim fila As Long = repeat_fila(user, "Usuarios", 2, 1, 0)
        For i = 0 To botones.Length - 1
            botones(i).Enabled = workbook.Sheets("Usuarios").cells(fila, i + 6).value
        Next
        For i = 2 To 11
            workbook.Sheets(i).visible() = workbook.Sheets("Usuarios").cells(fila, i + 21).value
        Next
        Return Nothing
    End Function

    Public Function Numcaracteres(ByVal texto As String) As Integer
        Dim arraytext() = texto.ToCharArray()
        Return arraytext.Length
    End Function

    Public Function auto_codigo(ByVal text As String, ByVal hoja_partida As String, ByVal colum_partida As Integer)
        Dim fila = repeat_fila(text, "Codigos", 2, 1, 1)
        Dim codigo As String
        Dim referencias As Integer
        If fila <> 1 Then
            referencias = actualizar_cantidad(Convert.ToString(workbook.Sheets("Codigos").cells(fila, 2).value), hoja_partida, colum_partida) + 1
            If referencias < 10 Then
                codigo = workbook.Sheets("Codigos").cells(fila, 2).value & "-000" & referencias
            ElseIf referencias < 100 Then
                codigo = workbook.Sheets("Codigos").cells(fila, 2).value & "-00" & referencias
            ElseIf referencias < 1000 Then
                codigo = workbook.Sheets("Codigos").cells(fila, 2).value & "-0" & referencias
            Else
                codigo = workbook.Sheets("Codigos").cells(fila, 2).value & "-" & referencias
            End If
            Return codigo
        Else
            Return ""
        End If
    End Function

    Public Function actualizar_cantidad(ByVal texto As String, ByVal hoja As String, ByVal colum As Integer)
        Dim fila, num_mayor, cont_tam As Long
        Dim abreviatura As String
        Dim split(), vec() As String
        fila = recorrer_filas(1, hoja, 1)
        'for para hallar el valor del tamaño del vector
        For i = 1 To fila - 1
            abreviatura = workbook.Sheets(hoja).cells(i, colum).value
            split = abreviatura.Split("-")
            If split.Length > 1 Then
                If texto = split(0) Then
                    cont_tam = cont_tam + 1
                End If
            End If
        Next
        ReDim vec(cont_tam)
        Dim cont_ref As Long = 0
        For i = 1 To fila - 1
            abreviatura = workbook.Sheets(hoja).cells(i, colum).value
            split = abreviatura.Split("-")
            If split.Length > 1 Then
                If texto = split(0) Then
                    vec(cont_ref) = split(1)
                    cont_ref = cont_ref + 1
                End If
            End If
        Next
        Call ordenar_mayor_menor(vec)
        For i = 0 To vec.Length - 2
            If vec(i + 1) <> vec(i) Then
                If vec(i + 1) - vec(i) = 1 Then
                    num_mayor = vec(i + 1)
                Else
                    Exit For
                End If
            End If
        Next
        Return num_mayor
    End Function

    Public Function ordenar_mayor_menor(ByRef vec() As Object)
        Dim aux As Object
        For i = 0 To vec.Length - 1
            For j = 0 To vec.Length - 1
                If vec(i) < vec(j) Then
                    aux = vec(j)
                    vec(j) = vec(i)
                    vec(i) = aux
                End If
            Next
        Next
        Return Nothing
    End Function

    Public Function Guardar_Factura(ByVal hoja As String, ByRef codigo As String, ByRef Fecha As String, ByRef lista As ListView)
        Dim ubicacion As String
        Dim Compraoventa() As String
        Fecha = Replace(Fecha, "/", "-")
        Compraoventa = codigo.Split("-")
        If Compraoventa(0) = "VTA" Then
            ubicacion = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\QbitElectronica\Facturas Venta\"
        Else
            ubicacion = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\QbitElectronica\Facturas Compra\"
        End If
        My.Computer.FileSystem.CreateDirectory(ubicacion)
        workbook.Sheets(hoja).printOut(PrintToFile:=True, PrToFileName:=ubicacion & codigo & " " & Fecha & ".pdf", ActivePrinter:="Microsoft Print to PDF")
        ubicacion = ubicacion & codigo & " " & Fecha & ".pdf"
        Limpiar_Factura(lista)
        'MOSTRAR FACTURA AUN NO ARREGLADO
        'If Not My.Computer.FileSystem.DirectoryExists(ubicacion) Then
        'Diagnostics.Process.Start(ubicacion)
        'End If
        Return Nothing
    End Function

    Public Function Llenar_Factura(ByVal hoja As String, ByVal codigo As String, ByVal Fecha As String, ByVal nombre As String, ByVal id As Long,
                                   ByVal email As String, ByVal celular As Long, ByVal Descuento As String, ByRef Lista As ListView)
        Dim funcion = New Class_funciones
        Dim tot As Long = 0
        Dim f_list As Long = funcion.recorrer_filas(14, "Form. Facturas", 1)
        Dim iva As Integer = 19
        Dim consecutivo() As String = codigo.Split("-")
        Dim nom_ref() As String
        With workbook.Sheets(hoja)
            .Cells(1, 6).value = consecutivo(1)
            .Cells(3, 6).value = Fecha
            .Cells(10, 2).value = nombre
            .Cells(11, 5).value = id
            .Cells(11, 2).value = email
            .Cells(12, 5).value = celular
            .Cells(12, 2).value = "Bogotá D.C."
            .Cells(19, 6).value = Descuento
            .Cells(20, 6).value = iva
            .Cells(23, 2).value = "Los módulos electrónicos tiene garantía de 10 días después de su compra, y los componentes electrónicos solamente 5 días, todo ello por defectos de fabrica"
            .Cells(26, 3).value = "Contado-efectivo"
            .Cells(27, 3).value = "Inmediato"
            For i = 0 To Lista.Items.Count - 1
                nom_ref = Lista.Items(i).SubItems(2).Text.Split("-")
                .cells(f_list + i, 1).value = i + 1
                .cells(f_list + i, 2).value = Lista.Items(i).Text
                .cells(f_list + i, 3).value = nom_ref(0)
                .cells(f_list + i, 4).value = nom_ref(1)
                .cells(f_list + i, 5).value = Lista.Items(i).SubItems(4).Text
                .cells(f_list + i, 6).value = Lista.Items(i).Text * Lista.Items(i).SubItems(4).Text
                workbook.Sheets("Form. Facturas").Rows(f_list + i + 1).Insert()
                tot = tot + .cells(f_list + i, 6).value
                .cells(f_list + i + 4, 6).value = tot
                tot = tot * (iva / 100) + tot
                .cells(f_list + i + 7, 6).value = tot - (tot * (Descuento / 100))
            Next
        End With
        Return Nothing
    End Function
    Public Function Limpiar_Factura(ByRef Lista As ListView)
        For i = 0 To Lista.Items.Count - 1
            workbook.Sheets("Form. Facturas").Rows(16).Delete()
        Next
        For i = 0 To 5
            workbook.Sheets("Form. Facturas").cells(15, 1 + i).value = ""
        Next
        With workbook.Sheets("Form. Facturas")
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
End Class

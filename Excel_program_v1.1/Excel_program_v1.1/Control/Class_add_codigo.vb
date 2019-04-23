Public Class Class_add_codigo
    Public Function caja_en_blanco(ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef label As Label, ByVal hoja As String)
        Dim vector() As Object = {box1, box2}
        Dim funcion = New Class_funciones
        Dim estado As Byte = 0
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
            Call agregar_provee(box1, box2, label, hoja)
        End If
        Return Nothing
    End Function

    Public Function agregar_provee(ByRef prod As TextBox, ByRef abrev As TextBox, ByRef label1 As Label, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila, repetido, repeat As Long
        Dim texto As String
        'inicializar rango para agregar el producto
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        'enviar fila para buqueda de espacio en blanco
        fila = funcion.recorrer_filas(fila, hoja, 1)
        'verificar que esten repetidos
        repetido = funcion.repeat_fila(abrev.Text, hoja, 1, 2, 0)
        repeat = funcion.repeat_fila(prod.Text, hoja, 1, 1, 0)
        texto = workbook.Sheets(hoja).cells(repetido + 1, 1).value
        If repetido = fila And repeat = fila Then
            'llenar datos en la fila con el espacio en blanco
            With workbook.Sheets(hoja)
                .Cells(fila, 1).value = prod.Text
                .Cells(fila, 2).value = abrev.Text
                .cells(fila, 3).value = ribbon1.Button4.Label
            End With
            'seleccionar la fila donde se agrego el nuevo producto
            workbook.Cells(fila, 1).select()
            label1.Text = "Registrado"
            Call limpiar_box(prod, abrev)
        Else
            label1.Text = "El producto o abreviatura ya existe"
            'boton para modificar
            Dim mod_prod = New Class_edit_codigo
            Dim frm_edit_codigo = New Form_edit_codigo
            Dim frm_conf_editprod = New Form_conf_edi_prod
            ' condicion para abrir form de modificar producto
            If mod_prod.repeat_open(frm_conf_editprod) = True Then
                frm_edit_codigo.ComboBox1.Text = texto
                frm_edit_codigo.ShowDialog()
                Call limpiar_box(abrev, prod)
            End If
        End If
        Return Nothing
    End Function

    Public Function limpiar_box(ByRef abrev As TextBox, ByRef prod As TextBox)
        abrev.Text = ""
        prod.Text = ""
        Return Nothing
    End Function
End Class

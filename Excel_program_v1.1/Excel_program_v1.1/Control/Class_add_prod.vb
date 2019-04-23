Public Class Class_add_prod
    Dim funcion = New Class_funciones

    Public Function caja_en_blanco(ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As TextBox, ByRef richbox1 As RichTextBox,
                                   ByRef box5 As TextBox, ByRef box6 As TextBox, ByRef box7 As TextBox, ByRef label As Label)
        Dim vector() As Object = {box1, box2, box3, box5, box6, box7}
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
            'condicion para que el precio de venta sea mayor al de compra
            If funcion.precio_coherente(box7.Text, box6.Text) = False Then
                label.Text = "El valor de compra no puede ser mayor al de venta"
                box6.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                box7.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                estado = 1
            End If
        End If
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call registar_prod(box1, box2, box3, richbox1, box5, box6, box7, label, "Productos")
        End If
        Return Nothing
    End Function

    Public Function registar_prod(ByRef cod As TextBox, ByRef nom As TextBox, ByRef ref As TextBox,
                                  ByRef des As RichTextBox, ByRef fab As TextBox, ByRef cosu As TextBox,
                                  ByRef pres As TextBox, ByRef label1 As Label, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila, rep As Long
        Dim texto As String
        'inicializar rango para agregar el producto
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        'enviar fila para buqueda de espacio en blanco
        fila = funcion.recorrer_filas(fila, hoja, 1)
        'verificar que esten repetidos
        rep = funcion.repetir_valor(fila, hoja, cod.Text, nom.Text, ref.Text, fab.Text, 1, 2, 3, 5)
        texto = workbook.Sheets(hoja).cells(rep + 1, 1).value
        If rep = fila Then
            'llenar datos en la fila con el espacio en blanco
            With workbook.Sheets(hoja)
                .Cells(fila, 1).value = cod.Text
                .Cells(fila, 2).value = nom.Text
                .Cells(fila, 3).value = ref.Text
                .Cells(fila, 4).value = des.Text
                .Cells(fila, 5).value = fab.Text
                .Cells(fila, 6).value = cosu.Text
                .Cells(fila, 7).value = pres.Text
                .Cells(fila, 8).value = ribbon1.Button4.Label
            End With
            'seleccionar la fila donde se agrego el nuevo producto
            workbook.Cells(fila, 1).select()
            label1.Text = "Producto registrado"
            Call limpiar_box(cod, nom, ref, des, fab, cosu, pres)
        Else
            workbook.Sheets(hoja).Range(workbook.Sheets(hoja).cells(rep + 1, 1), workbook.Sheets(hoja).cells(rep + 1, 8)).select()
            label1.Text = "El producto o codigo ya existe"
            'boton para modificar
            Dim mod_prod = New Class_edit__prod
            Dim frm_edit_prod = New Form_edit_prod
            Dim frm_conf_editprod = New Form_conf_edi_prod
            ' condicion para habrir form de modificar producto
            If mod_prod.repeat_open(frm_conf_editprod) = True Then
                frm_edit_prod.ComboBox1.Text = texto
                frm_edit_prod.ShowDialog()
                Call limpiar_box(cod, nom, ref, des, fab, cosu, pres)
            End If
        End If
        Return Nothing
    End Function

    Public Function limpiar_box(ByRef cod As TextBox, ByRef nom As TextBox, ByRef ref As TextBox, ByRef des As RichTextBox, ByRef fab As TextBox,
                                ByRef cosu As TextBox, ByRef pres As TextBox)
        nom.Text = ""
        cod.Text = ""
        ref.Text = ""
        des.Text = ""
        fab.Text = ""
        cosu.Text = "0"
        pres.Text = "0"
        Return Nothing
    End Function
End Class

Public Class Class_add_client
    Dim funcion = New Class_funciones
    Public Function caja_en_blanco(ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As TextBox,
                                   ByRef box4 As TextBox, ByRef box5 As TextBox, ByRef label As Label,
                                   ByVal hoja As String)
        Dim vector() As Object = {box2, box4, box1, box3, box5}
        Dim estado As Byte = 0
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                estado = 1
            ElseIf i > 1 And funcion.es_numero(vector(i).Text) = False Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "Los valores asignados no son numeros"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call agregar_client(box1, box2, box3, box4, box5, label, hoja)
        End If
        Return Nothing
    End Function

    Public Function agregar_client(ByRef id As TextBox, ByRef nom As TextBox, ByRef celular As TextBox,
                                   ByRef email As TextBox, ByRef descuento As TextBox, ByRef label1 As Label,
                                   ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila, repetido As Long
        Dim texto As String
        'inicializar rango para agregar el producto
        workbook.Range("A1").Select()
        fila = workbook.ActiveCell.Row
        'enviar fila para buqueda de espacio en blanco
        fila = funcion.recorrer_filas(fila, hoja, 1)
        'verificar que esten repetidos
        repetido = funcion.repetir_valor(fila, hoja, id.Text, celular.Text, nom.Text, email.Text, 1, 3, 1, 4)
        texto = workbook.Sheets(hoja).cells(repetido + 1, 1).value
        If repetido = fila Then
            'llenar datos en la fila con el espacio en blanco
            With workbook.Sheets(hoja)
                .Cells(fila, 1).value = id.Text
                .Cells(fila, 2).value = nom.Text
                .Cells(fila, 3).value = celular.Text
                .Cells(fila, 4).value = email.Text
                .Cells(fila, 5).value = descuento.Text
                .Cells(fila, 6).value = ribbon1.Button4.Label
            End With
            'seleccionar la fila donde se agrego el nuevo producto
            workbook.Cells(fila, 1).select()
            label1.Text = "Registrado"
            Call limpiar_box(id, nom, celular, email, descuento)
        Else
            workbook.Sheets(hoja).Range(workbook.Sheets(hoja).cells(repetido + 1, 1), workbook.Sheets(hoja).cells(repetido + 1, 4)).select()
            label1.Text = "El producto o codigo ya existe"
            'boton para modificar
            Dim mod_prod = New Class_edit_client
            Dim frm_edit_client = New Form_edit_client
            Dim frm_conf_editprod = New Form_conf_edi_prod
            ' condicion para abrir form de modificar producto
            If mod_prod.repeat_open(frm_conf_editprod) = True Then
                frm_edit_client.ComboBox1.Text = texto
                frm_edit_client.ShowDialog()
                Call limpiar_box(id, nom, celular, email, descuento)
            End If
        End If
        Return Nothing
    End Function

    Public Function limpiar_box(ByRef id As TextBox, ByRef nom As TextBox, ByRef celular As TextBox,
                                ByRef email As TextBox, ByRef descuento As TextBox)
        nom.Text = ""
        celular.Text = ""
        id.Text = ""
        email.Text = ""
        descuento.Text = "0"
        Return Nothing
    End Function
End Class

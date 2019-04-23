Public Class Class_add_provee
    Dim funcion = New Class_funciones
    Public Function caja_en_blanco(ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As TextBox, ByRef box4 As TextBox,
                                   ByRef box5 As TextBox, ByRef label As Label)
        Dim vector() As Object = {box1, box2, box5, box3, box4}
        Dim estado As Byte = 0
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Or vector(i).text = "0" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                estado = 1
            ElseIf i > 2 And funcion.es_numero(vector(i).Text) = False Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "Los valores asignados no son numeros"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call agregar_provee(box1, box2, box3, box4, box5, label)
        End If
        Return Nothing
    End Function

    Public Function agregar_provee(ByRef nom As TextBox, ByRef repre As TextBox, ByRef celular As TextBox, ByRef nit As TextBox,
                                   ByRef email As TextBox, ByRef label1 As Label)
        Dim funcion = New Class_funciones
        Dim workbook = Globals.ThisWorkbook.Application
        Dim ribbon1 = Globals.Ribbons.Ribbon1
        Dim fila, repetido, rep_nit As Long
        Dim texto As String
        'enviar fila para buqueda de espacio en blanco
        fila = funcion.recorrer_filas(1, "Proveedores", 1)
        'verificar que esten repetidos
        repetido = funcion.repetir_valor(fila, "Proveedores", nom.Text, repre.Text, celular.Text, email.Text, 1, 2, 3, 5)
        rep_nit = funcion.repeat_fila(nit.Text, "Proveedores", 4, 1, 1)
        texto = workbook.Sheets("Proveedores").cells(repetido + 1, 1).value
        If repetido = fila And rep_nit = 1 Then
            'llenar datos en la fila con el espacio en blanco
            With workbook.Sheets("Proveedores")
                .Cells(fila, 1).value = nom.Text
                .Cells(fila, 2).value = repre.Text
                .Cells(fila, 3).value = celular.Text
                .Cells(fila, 4).value = nit.Text
                .Cells(fila, 5).value = email.Text
                .Cells(fila, 6).value = ribbon1.Button4.Label
            End With
            'seleccionar la fila donde se agrego el nuevo producto
            workbook.Cells(fila, 1).select()
            label1.Text = "Registrado"
            Call limpiar_box(nom, repre, celular, nit, email)
        Else
            workbook.Sheets("Proveedores").Range(workbook.Sheets("Proveedores").cells(repetido + 1, 1), workbook.Sheets("Proveedores").cells(repetido + 1, 5)).select()
            label1.Text = "El producto o codigo ya existe"
            'boton para modificar
            Dim mod_prod = New Class_edit_provee
            Dim frm_conf_editprod = New Form_conf_edi_prod
            Dim frm_edit_provee = New Form_edit_provee
            ' condicion para abrir form de modificar producto
            If mod_prod.repeat_open(frm_conf_editprod) = True Then
                label1.Text = "Modificando producto existente"
                frm_edit_provee.ComboBox1.Text = texto
                frm_edit_provee.ShowDialog()
                Call limpiar_box(nom, repre, celular, nit, email)
                label1.Text = "Registrando..."
            End If
        End If

        Return Nothing
    End Function

    Public Function limpiar_box(ByRef nom As TextBox, ByRef repre As TextBox, ByRef celular As TextBox, ByRef nit As TextBox, ByRef email As TextBox)
        nom.Text = ""
        repre.Text = ""
        celular.Text = ""
        nit.Text = ""
        email.Text = ""
        Return Nothing
    End Function
End Class

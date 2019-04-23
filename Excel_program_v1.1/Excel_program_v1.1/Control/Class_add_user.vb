Public Class Class_add_user
    Public Function caja_en_blanco(ByRef box1 As TextBox, ByRef box2 As TextBox, ByRef box3 As TextBox, ByRef box4 As TextBox,
                                   ByRef box5 As TextBox, ByRef checkbox1 As CheckBox, ByRef checkbox2 As CheckBox, ByRef label As Label,
                                   ByVal hoja As String)
        Dim vector() As Object = {box2, box4, box1, box3, box5}
        Dim estado As Byte = 0
        Dim funcion = New Class_funciones
        'condiciones que evaluan los text box vacios y demarcan en color rojo
        For i = 0 To vector.Length - 1
            If vector(i).text = "" Then
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
                label.Text = "No pueden haber espacios en blanco"
                estado = 1
            Else
                vector(i).BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            End If
        Next
        If funcion.Numcaracteres(box2.Text) < 8 Then
            box2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label.Text = "No puede tener menos de 8 caracteres"
            estado = 1
        ElseIf box2.Text <> box3.Text Then
            box3.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            box2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            label.Text = "Las contraseñas no coinciden"
            estado = 1
        End If

        If checkbox1.CheckState = 0 And checkbox2.CheckState = 0 Then
            estado = 1
            label.Text = "Los checkboxs no pueden estar vacios"
            checkbox1.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            checkbox2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
        ElseIf checkbox1.CheckState = 1 And checkbox2.CheckState = 1 Then
            estado = 1
            label.Text = "Los checkboxs no pueden estar precionados al tiempo"
            checkbox1.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
            checkbox2.BackColor = Drawing.ColorTranslator.FromHtml("#FF8585")
        Else
            checkbox1.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            checkbox2.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        ' condiciones para que el contador autorize el registro
        If estado = 0 Then
            Call agregar_user(box1, box2, box3, box4, box5, checkbox1, checkbox2, label, hoja)
        End If
        Return Nothing
    End Function

    Public Function agregar_user(ByRef user As TextBox, ByRef contra As TextBox, ByRef repeat_contra As TextBox,
                                 ByRef nomb As TextBox, ByRef correo As TextBox, ByRef checkbox1 As CheckBox,
                                 ByRef checkbox2 As CheckBox, ByRef label1 As Label, ByVal hoja As String)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim funcion = New Class_funciones
        Dim fila, repetido As Long
        Dim texto As String
        fila = funcion.recorrer_filas(1, hoja, 1)
        repetido = funcion.repeat_fila(user.Text, hoja, 1, 1, 0)
        texto = workbook.Sheets(hoja).cells(repetido + 1, 1).value
        If repetido = fila Then
            'llenar datos en la fila con el espacio en blanco
            With workbook.Sheets(hoja)
                .Cells(fila, 1).value = user.Text
                .Cells(fila, 2).value = contra.Text
                .Cells(fila, 3).value = nomb.Text
                .Cells(fila, 4).value = correo.Text
            End With
            'añadir user o admin
            If checkbox1.CheckState = 1 Then
                workbook.Sheets(hoja).cells(fila, 5).value = "Trabajador"
            Else
                workbook.Sheets(hoja).cells(fila, 5).value = "Administrador"
            End If
            tipodeusuario(workbook.Sheets(hoja).cells(fila, 5).value, fila)
            'seleccionar la fila donde se agrego el nuevo producto
            workbook.Cells(fila, 1).select()
            label1.Text = "Registrado"
            Call limpiar_box(user, contra, repeat_contra, nomb, correo, checkbox1, checkbox2)
        Else
            workbook.Sheets(hoja).Range(workbook.Sheets(hoja).cells(repetido + 1, 1), workbook.Sheets(hoja).cells(repetido + 1, 4)).select()
            label1.Text = "El producto o codigo ya existe"
        End If
        Return Nothing
    End Function

    Public Function tipodeusuario(ByVal tipo_usuario As String, ByVal fila As String)
        Dim workbook = Globals.ThisWorkbook.Application
        If tipo_usuario = "Administrador" Then
            For i = 6 To 32
                If i = 6 Or i = 29 Or i = 30 Or i = 31 Then
                    workbook.Sheets("Usuarios").cells(fila, i).value = False
                Else
                    workbook.Sheets("Usuarios").cells(fila, i).value = True
                End If
            Next
        ElseIf tipo_usuario = "Trabajador" Then
            For i = 6 To 32
                If i < 8 Or i = 9 Or i = 20 Or i = 21 Or i = 22 Or i > 27 Then
                    workbook.Sheets("Usuarios").cells(fila, i).value = False
                Else
                    workbook.Sheets("Usuarios").cells(fila, i).value = True
                End If
            Next
        End If
        Return Nothing
    End Function

    Public Function limpiar_box(ByRef user As TextBox, ByRef contra As TextBox, ByRef repeat_contra As TextBox,
                                ByRef nomb As TextBox, ByRef correo As TextBox, ByRef checkbox1 As CheckBox,
                                ByRef checkbox2 As CheckBox)
        user.Text = ""
        contra.Text = ""
        repeat_contra.Text = ""
        nomb.Text = ""
        correo.Text = ""
        checkbox1.CheckState = False
        checkbox1.CheckState = False
        Return Nothing
    End Function
End Class
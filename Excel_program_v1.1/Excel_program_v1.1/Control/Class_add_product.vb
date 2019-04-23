Public Class Class_add_product
    Dim funcion = New Class_funciones
    Public Function espacios_vacios(ByRef cod As TextBox, ByRef nom As TextBox, ByRef ref As TextBox,
                                    ByRef desc As TextBox, ByRef fab As TextBox, ByRef cost As TextBox,
                                    ByRef precio As TextBox, ByRef labelvar As Label)
        Dim vacio As Integer = 0
        If nom.Text = "" Then
            nom.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF8585")
            labelvar.Text = "No pueden haber casillas vacias"
            vacio = 1
        Else
            nom.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        If ref.Text = "" Then
            ref.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF8585")
            labelvar.Text = "No pueden haber casillas vacias"
            vacio = 1
        Else
            ref.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        If fab.Text = "" Then
            fab.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF8585")
            labelvar.Text = "No pueden haber casillas vacias"
            vacio = 1
        Else
            fab.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        'Comprobación letras en precio o costo
        If cost.Text = 0 And funcion.comp_numeros(cost.Text) = False Then
            cost.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF8585")
            labelvar.Text = "Los datos de costo y/o precio no pueden ser letras"
            vacio = 1
        Else
            cost.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        If precio.Text = 0 And funcion.comp_numeros(precio.Text) = False Then
            precio.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF8585")
            labelvar.Text = "Los datos de costo y/o precio no pueden ser letras"
            vacio = 1
        Else
            precio.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End If
        'comprobacion de que el formulacion se halla llenado correctamente 
        If vacio = 0 Then
            labelvar.Text = "Esperando registro del nuevo producto"
            registrar_product(cod, nom, ref, desc, fab, cost, precio, labelvar)
        End If
        Return Nothing
    End Function

    Public Function registrar_product(ByRef cod As TextBox, ByRef nom As TextBox, ByRef ref As TextBox,
                                      ByRef desc As TextBox, ByRef fab As TextBox, ByRef cost As TextBox,
                                      ByRef precio As TextBox, ByRef labelvar As Label)
        Dim workbook = Globals.ThisWorkbook.Application
        Dim edit_pro = New Class_edit_product
        Dim modi As Integer
        Dim fila As Long = funcion.Add_registro(2, 2)
        Dim antirepetido As Integer
        antirepetido = funcion.Buscar_registro(2, 2, nom.Text, ref.Text, fab.Text, fila)
        If antirepetido = fila Then
            With workbook.Sheets(2)
                .Cells(fila, 1).value = cod.Text
                .Cells(fila, 2).value = nom.Text
                .Cells(fila, 3).value = ref.Text
                .Cells(fila, 4).value = desc.Text
                .Cells(fila, 5).value = fab.Text
                .Cells(fila, 6).value = cost.Text
                .Cells(fila, 7).value = precio.Text
            End With
            labelvar.Text = "El producto ha sido registrado exitosamente"
            limpiar_repeat(cod, nom, ref, desc, fab, cost, precio)
        Else
            workbook.Sheets(2).Range(workbook.Sheets(2).cells(antirepetido + 1, 1), workbook.Sheets(2).cells(antirepetido + 1, 1)).select()
            modi = funcion.repeat_question("¿DESEA MODIFICAR EL PRODUCTO?")
            labelvar.Text = "El producto que ingreso ya existe"
            If modi = 1 Then
                edit_pro.repeat_abrir()
                limpiar_repeat(cod, nom, ref, desc, fab, cost, precio)
                labelvar.Text = "El producto ha sido modificado exitosamente"
            End If
        End If
        Return Nothing
    End Function

    Public Function limpiar_repeat(ByRef cod As TextBox, ByRef nom As TextBox, ByRef ref As TextBox,
                                   ByRef desc As TextBox, ByRef fab As TextBox, ByRef cost As TextBox,
                                   ByRef precio As TextBox)
        cod.Text = ""
        nom.Text = ""
        ref.Text = ""
        desc.Text = ""
        fab.Text = ""
        cost.Text = "0"
        precio.Text = "0"
        Return Nothing
    End Function
End Class

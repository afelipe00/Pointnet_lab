Imports Microsoft.Office.Tools.Ribbon
Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim frm_ajustes = New Form_ajustes
        frm_ajustes.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        Dim frm_adduser = New Form_add_user
        frm_adduser.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        Dim frm_removeuser = New Form_edit_user
        frm_removeuser.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        Dim frm_permisosyres = New Form_permisos
        frm_permisosyres.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Productos").Select()
        Dim frm_addproducto = New Form_add_prod
        frm_addproducto.ShowDialog()
        Workbook.Sheets("Principal").Select()

    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Productos").Select()
        Dim frm_editarprod = New Form_edit_prod
        frm_editarprod.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Proveedores").Select()
        Dim frm_agregarprovee = New Form_add_provee
        frm_agregarprovee.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Proveedores").Select()
        Dim frm_eliminarprovee = New Form_edit_provee
        frm_eliminarprovee.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Mov. Inventario").Select()
        Dim frm_agregarcompra = New Form_add_purchase
        frm_agregarcompra.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Mov. Inventario").Select()
        Dim frm_devcompra = New Form_return_purchase
        frm_devcompra.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets(4).Select()
        Dim frm_ventas = New Form_sales
        frm_ventas.Show()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets(4).Select()
        Dim frm_devventas = New Form_return_sales
        frm_devventas.Show()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Clientes").Select()
        Dim frm_addcliente = New Form_add_client
        frm_addcliente.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        Dim Workbook = Globals.ThisWorkbook.Application
        Workbook.Sheets("Clientes").Select()
        Dim frm_editarcliente = New Form_edit_client
        frm_editarcliente.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        Dim frm_movxproducto = New Form_mov_product
        Dim Workbook = Globals.ThisWorkbook.Application
        frm_movxproducto.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        Dim frm_movxfecha = New Form_mov_fecha
        Dim Workbook = Globals.ThisWorkbook.Application
        frm_movxfecha.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        Dim frm_existencia = New Form_existencias
        Dim Workbook = Globals.ThisWorkbook.Application
        frm_existencia.ShowDialog()
        Workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        Dim frm_compcompra = New Form_comp_purchase
        frm_compcompra.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        Dim frm_compventa = New Form_comp_sale
        frm_compventa.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        Dim frm_cotizacion = New Form_cotizacion
        frm_cotizacion.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        Dim frm_libromayor = New Form_book_up
        frm_libromayor.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        Dim frm_estadoresultados = New Form_estado_resultados
        frm_estadoresultados.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        Dim frm_balancegen = New Form_balance_general
        frm_balancegen.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim frm_login = New Form_Login_other
        frm_login.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click
        Dim frm_estado = New Form_estado_devolucion_compra
        frm_estado.ShowDialog()
        Dim workbook = Globals.ThisWorkbook.Application
        workbook.Sheets("Principal").Select()
    End Sub
End Class

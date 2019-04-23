Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Requerido para la compatibilidad con el Diseñador de composiciones de clases Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'El Diseñador de componentes requiere esta llamada.
        InitializeComponent()

    End Sub

    'Component reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Menu4 = Me.Factory.CreateRibbonMenu
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Menu5 = Me.Factory.CreateRibbonMenu
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Label = "Administrador de inventario"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button4)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Menu4)
        Me.Group1.Items.Add(Me.Menu1)
        Me.Group1.Label = "Perfil"
        Me.Group1.Name = "Group1"
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "Actual"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "Ajustes"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Menu4
        '
        Me.Menu4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu4.Image = CType(resources.GetObject("Menu4.Image"), System.Drawing.Image)
        Me.Menu4.Items.Add(Me.Button10)
        Me.Menu4.Items.Add(Me.Button11)
        Me.Menu4.Items.Add(Me.Button12)
        Me.Menu4.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu4.Label = "Usuarios"
        Me.Menu4.Name = "Menu4"
        Me.Menu4.ShowImage = True
        '
        'Button10
        '
        Me.Button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button10.Image = CType(resources.GetObject("Button10.Image"), System.Drawing.Image)
        Me.Button10.Label = "Agregar usuario"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button11
        '
        Me.Button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.Label = "Editar usuario"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Button12
        '
        Me.Button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button12.Image = CType(resources.GetObject("Button12.Image"), System.Drawing.Image)
        Me.Button12.Label = "Permisos y resticciones"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Image = CType(resources.GetObject("Menu1.Image"), System.Drawing.Image)
        Me.Menu1.Items.Add(Me.Button2)
        Me.Menu1.Items.Add(Me.Button3)
        Me.Menu1.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Label = "Productos"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.ShowImage = True
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "Agregar producto"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "Editar producto"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Menu2)
        Me.Group2.Items.Add(Me.Menu3)
        Me.Group2.Label = "Entradas"
        Me.Group2.Name = "Group2"
        '
        'Menu2
        '
        Me.Menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu2.Dynamic = True
        Me.Menu2.Image = CType(resources.GetObject("Menu2.Image"), System.Drawing.Image)
        Me.Menu2.Items.Add(Me.Button13)
        Me.Menu2.Items.Add(Me.Button14)
        Me.Menu2.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu2.Label = "Proveedores"
        Me.Menu2.Name = "Menu2"
        Me.Menu2.ShowImage = True
        '
        'Button13
        '
        Me.Button13.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button13.Image = CType(resources.GetObject("Button13.Image"), System.Drawing.Image)
        Me.Button13.Label = "Agregar proveedor"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        '
        'Button14
        '
        Me.Button14.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button14.Image = CType(resources.GetObject("Button14.Image"), System.Drawing.Image)
        Me.Button14.Label = "Editar proveedor"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Menu3
        '
        Me.Menu3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu3.Image = CType(resources.GetObject("Menu3.Image"), System.Drawing.Image)
        Me.Menu3.Items.Add(Me.Button8)
        Me.Menu3.Items.Add(Me.Button9)
        Me.Menu3.Items.Add(Me.Button18)
        Me.Menu3.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu3.Label = "Compras"
        Me.Menu3.Name = "Menu3"
        Me.Menu3.ShowImage = True
        '
        'Button8
        '
        Me.Button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button8.Image = CType(resources.GetObject("Button8.Image"), System.Drawing.Image)
        Me.Button8.Label = "Agregar Compra"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        '
        'Button9
        '
        Me.Button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button9.Image = CType(resources.GetObject("Button9.Image"), System.Drawing.Image)
        Me.Button9.Label = "Devolver compra"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button18
        '
        Me.Button18.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button18.Image = CType(resources.GetObject("Button18.Image"), System.Drawing.Image)
        Me.Button18.Label = "Estado devolución"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button5)
        Me.Group3.Items.Add(Me.Menu5)
        Me.Group3.Items.Add(Me.Button6)
        Me.Group3.Label = "Salidas"
        Me.Group3.Name = "Group3"
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "Ventas"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Menu5
        '
        Me.Menu5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu5.Image = CType(resources.GetObject("Menu5.Image"), System.Drawing.Image)
        Me.Menu5.Items.Add(Me.Button7)
        Me.Menu5.Items.Add(Me.Button19)
        Me.Menu5.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu5.Label = "Clientes"
        Me.Menu5.Name = "Menu5"
        Me.Menu5.ShowImage = True
        '
        'Button7
        '
        Me.Button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.Label = "Agregar cliente"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button19
        '
        Me.Button19.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button19.Image = CType(resources.GetObject("Button19.Image"), System.Drawing.Image)
        Me.Button19.Label = "Editar cliente"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Button6
        '
        Me.Button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.Label = "Devolucion ventas"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button16)
        Me.Group4.Items.Add(Me.Button17)
        Me.Group4.Items.Add(Me.Button15)
        Me.Group4.Label = "Consultas"
        Me.Group4.Name = "Group4"
        '
        'Button16
        '
        Me.Button16.Image = CType(resources.GetObject("Button16.Image"), System.Drawing.Image)
        Me.Button16.Label = "Movimientos por producto"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'Button17
        '
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.Label = "Movimientos por fecha"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button15
        '
        Me.Button15.Image = CType(resources.GetObject("Button15.Image"), System.Drawing.Image)
        Me.Button15.Label = "Productos en existencia"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button22)
        Me.Group5.Items.Add(Me.Button21)
        Me.Group5.Items.Add(Me.Button20)
        Me.Group5.Label = "Documentos"
        Me.Group5.Name = "Group5"
        '
        'Button22
        '
        Me.Button22.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button22.Image = CType(resources.GetObject("Button22.Image"), System.Drawing.Image)
        Me.Button22.Label = "Comprobante de compra"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        '
        'Button21
        '
        Me.Button21.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button21.Image = CType(resources.GetObject("Button21.Image"), System.Drawing.Image)
        Me.Button21.Label = "Comprobante de venta"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button20
        '
        Me.Button20.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button20.Image = CType(resources.GetObject("Button20.Image"), System.Drawing.Image)
        Me.Button20.Label = "Cotización"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.Button23)
        Me.Group6.Items.Add(Me.Button24)
        Me.Group6.Items.Add(Me.Button25)
        Me.Group6.Label = "Informes"
        Me.Group6.Name = "Group6"
        '
        'Button23
        '
        Me.Button23.Image = CType(resources.GetObject("Button23.Image"), System.Drawing.Image)
        Me.Button23.Label = "Libro mayor"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        '
        'Button24
        '
        Me.Button24.Image = CType(resources.GetObject("Button24.Image"), System.Drawing.Image)
        Me.Button24.Label = "Estado de resultados"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'Button25
        '
        Me.Button25.Image = CType(resources.GetObject("Button25.Image"), System.Drawing.Image)
        Me.Button25.Label = "Balance general"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu5 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class

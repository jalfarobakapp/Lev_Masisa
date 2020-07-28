<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Levantar_Tablas
    Inherits DevComponents.DotNetBar.Metro.MetroForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Levantar_Tablas))
        Me.Grupo_01 = New DevComponents.DotNetBar.Controls.GroupPanel()
        Me.Chk_Primera_Fila_Es_encabezado = New DevComponents.DotNetBar.Controls.CheckBoxX()
        Me.Lbl_Procesando = New DevComponents.DotNetBar.LabelX()
        Me.Txt_Nombre_Archivo = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.Circular_Progres_Porc = New DevComponents.DotNetBar.Controls.CircularProgress()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Circular_Progres_Val = New DevComponents.DotNetBar.Controls.CircularProgress()
        Me.Bar1 = New DevComponents.DotNetBar.Bar()
        Me.Btn_Buscar_Archivo = New DevComponents.DotNetBar.ButtonItem()
        Me.Btn_Importar_Desde_XML = New DevComponents.DotNetBar.ButtonItem()
        Me.Btn_Archivo_Ayuda_Excel = New DevComponents.DotNetBar.ButtonItem()
        Me.Btn_Cancelar = New DevComponents.DotNetBar.ButtonItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Grupo_01.SuspendLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Grupo_01
        '
        Me.Grupo_01.BackColor = System.Drawing.Color.White
        Me.Grupo_01.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Grupo_01.Controls.Add(Me.Chk_Primera_Fila_Es_encabezado)
        Me.Grupo_01.Controls.Add(Me.Lbl_Procesando)
        Me.Grupo_01.Controls.Add(Me.Txt_Nombre_Archivo)
        Me.Grupo_01.Controls.Add(Me.Circular_Progres_Porc)
        Me.Grupo_01.Controls.Add(Me.Label1)
        Me.Grupo_01.Controls.Add(Me.Circular_Progres_Val)
        Me.Grupo_01.DisabledBackColor = System.Drawing.Color.Empty
        Me.Grupo_01.Location = New System.Drawing.Point(12, 12)
        Me.Grupo_01.Name = "Grupo_01"
        Me.Grupo_01.Size = New System.Drawing.Size(585, 140)
        '
        '
        '
        Me.Grupo_01.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2
        Me.Grupo_01.Style.BackColorGradientAngle = 90
        Me.Grupo_01.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground
        Me.Grupo_01.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.Grupo_01.Style.BorderBottomWidth = 1
        Me.Grupo_01.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder
        Me.Grupo_01.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.Grupo_01.Style.BorderLeftWidth = 1
        Me.Grupo_01.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.Grupo_01.Style.BorderRightWidth = 1
        Me.Grupo_01.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.Grupo_01.Style.BorderTopWidth = 1
        Me.Grupo_01.Style.CornerDiameter = 4
        Me.Grupo_01.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded
        Me.Grupo_01.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center
        Me.Grupo_01.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText
        Me.Grupo_01.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near
        '
        '
        '
        Me.Grupo_01.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.Grupo_01.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Grupo_01.TabIndex = 80
        Me.Grupo_01.Text = "Importar datos desde Excel"
        '
        'Chk_Primera_Fila_Es_encabezado
        '
        Me.Chk_Primera_Fila_Es_encabezado.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.Chk_Primera_Fila_Es_encabezado.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Chk_Primera_Fila_Es_encabezado.Checked = True
        Me.Chk_Primera_Fila_Es_encabezado.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Chk_Primera_Fila_Es_encabezado.CheckValue = "Y"
        Me.Chk_Primera_Fila_Es_encabezado.ForeColor = System.Drawing.Color.Black
        Me.Chk_Primera_Fila_Es_encabezado.Location = New System.Drawing.Point(3, 89)
        Me.Chk_Primera_Fila_Es_encabezado.Name = "Chk_Primera_Fila_Es_encabezado"
        Me.Chk_Primera_Fila_Es_encabezado.Size = New System.Drawing.Size(286, 23)
        Me.Chk_Primera_Fila_Es_encabezado.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.Chk_Primera_Fila_Es_encabezado.TabIndex = 81
        Me.Chk_Primera_Fila_Es_encabezado.Text = "La primera fila es el encabezado"
        '
        'Lbl_Procesando
        '
        Me.Lbl_Procesando.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.Lbl_Procesando.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Lbl_Procesando.ForeColor = System.Drawing.Color.Black
        Me.Lbl_Procesando.Location = New System.Drawing.Point(5, 64)
        Me.Lbl_Procesando.Name = "Lbl_Procesando"
        Me.Lbl_Procesando.Size = New System.Drawing.Size(571, 19)
        Me.Lbl_Procesando.TabIndex = 78
        Me.Lbl_Procesando.Text = "..."
        '
        'Txt_Nombre_Archivo
        '
        Me.Txt_Nombre_Archivo.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.Txt_Nombre_Archivo.Border.Class = "TextBoxBorder"
        Me.Txt_Nombre_Archivo.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Txt_Nombre_Archivo.DisabledBackColor = System.Drawing.Color.White
        Me.Txt_Nombre_Archivo.ForeColor = System.Drawing.Color.Black
        Me.Txt_Nombre_Archivo.Location = New System.Drawing.Point(118, 30)
        Me.Txt_Nombre_Archivo.Name = "Txt_Nombre_Archivo"
        Me.Txt_Nombre_Archivo.PreventEnterBeep = True
        Me.Txt_Nombre_Archivo.ReadOnly = True
        Me.Txt_Nombre_Archivo.Size = New System.Drawing.Size(458, 22)
        Me.Txt_Nombre_Archivo.TabIndex = 8
        '
        'Circular_Progres_Porc
        '
        Me.Circular_Progres_Porc.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.Circular_Progres_Porc.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Circular_Progres_Porc.Location = New System.Drawing.Point(63, 14)
        Me.Circular_Progres_Porc.Name = "Circular_Progres_Porc"
        Me.Circular_Progres_Porc.ProgressBarType = DevComponents.DotNetBar.eCircularProgressType.Donut
        Me.Circular_Progres_Porc.ProgressColor = System.Drawing.Color.SteelBlue
        Me.Circular_Progres_Porc.ProgressTextVisible = True
        Me.Circular_Progres_Porc.Size = New System.Drawing.Size(49, 44)
        Me.Circular_Progres_Porc.Style = DevComponents.DotNetBar.eDotNetBarStyle.OfficeXP
        Me.Circular_Progres_Porc.TabIndex = 77
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(115, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Nombre de archivo"
        '
        'Circular_Progres_Val
        '
        Me.Circular_Progres_Val.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.Circular_Progres_Val.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.Circular_Progres_Val.Location = New System.Drawing.Point(3, 14)
        Me.Circular_Progres_Val.Name = "Circular_Progres_Val"
        Me.Circular_Progres_Val.ProgressBarType = DevComponents.DotNetBar.eCircularProgressType.Donut
        Me.Circular_Progres_Val.ProgressColor = System.Drawing.Color.SteelBlue
        Me.Circular_Progres_Val.ProgressTextFormat = "{0}"
        Me.Circular_Progres_Val.ProgressTextVisible = True
        Me.Circular_Progres_Val.Size = New System.Drawing.Size(54, 44)
        Me.Circular_Progres_Val.Style = DevComponents.DotNetBar.eDotNetBarStyle.OfficeXP
        Me.Circular_Progres_Val.TabIndex = 76
        '
        'Bar1
        '
        Me.Bar1.AntiAlias = True
        Me.Bar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Bar1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Bar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.Btn_Buscar_Archivo, Me.Btn_Importar_Desde_XML, Me.Btn_Archivo_Ayuda_Excel, Me.Btn_Cancelar})
        Me.Bar1.Location = New System.Drawing.Point(0, 165)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(608, 41)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.Bar1.TabIndex = 79
        Me.Bar1.TabStop = False
        Me.Bar1.Text = "Bar1"
        '
        'Btn_Buscar_Archivo
        '
        Me.Btn_Buscar_Archivo.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.Btn_Buscar_Archivo.ForeColor = System.Drawing.Color.Black
        Me.Btn_Buscar_Archivo.Image = CType(resources.GetObject("Btn_Buscar_Archivo.Image"), System.Drawing.Image)
        Me.Btn_Buscar_Archivo.Name = "Btn_Buscar_Archivo"
        Me.Btn_Buscar_Archivo.Text = "Importar archivo"
        Me.Btn_Buscar_Archivo.Tooltip = "Buscar Archivo"
        '
        'Btn_Importar_Desde_XML
        '
        Me.Btn_Importar_Desde_XML.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.Btn_Importar_Desde_XML.ForeColor = System.Drawing.Color.Black
        Me.Btn_Importar_Desde_XML.Image = CType(resources.GetObject("Btn_Importar_Desde_XML.Image"), System.Drawing.Image)
        Me.Btn_Importar_Desde_XML.Name = "Btn_Importar_Desde_XML"
        Me.Btn_Importar_Desde_XML.Text = "Importar archivo XML"
        Me.Btn_Importar_Desde_XML.Visible = False
        '
        'Btn_Archivo_Ayuda_Excel
        '
        Me.Btn_Archivo_Ayuda_Excel.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.Btn_Archivo_Ayuda_Excel.ForeColor = System.Drawing.Color.Black
        Me.Btn_Archivo_Ayuda_Excel.Image = CType(resources.GetObject("Btn_Archivo_Ayuda_Excel.Image"), System.Drawing.Image)
        Me.Btn_Archivo_Ayuda_Excel.Name = "Btn_Archivo_Ayuda_Excel"
        Me.Btn_Archivo_Ayuda_Excel.Tooltip = "Ayuda, ejemplo archivo excel."
        '
        'Btn_Cancelar
        '
        Me.Btn_Cancelar.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.Btn_Cancelar.ForeColor = System.Drawing.Color.Black
        Me.Btn_Cancelar.Image = CType(resources.GetObject("Btn_Cancelar.Image"), System.Drawing.Image)
        Me.Btn_Cancelar.Name = "Btn_Cancelar"
        Me.Btn_Cancelar.Text = "Cancelar"
        Me.Btn_Cancelar.Tooltip = "Eliminar Servidor de correo de salida SMTP"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Frm_Importar_Compras_SII
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(608, 206)
        Me.Controls.Add(Me.Grupo_01)
        Me.Controls.Add(Me.Bar1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Importar_Compras_SII"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Importar compras desde SII"
        Me.Grupo_01.ResumeLayout(False)
        Me.Grupo_01.PerformLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Grupo_01 As DevComponents.DotNetBar.Controls.GroupPanel
    Friend WithEvents Chk_Primera_Fila_Es_encabezado As DevComponents.DotNetBar.Controls.CheckBoxX
    Friend WithEvents Lbl_Procesando As DevComponents.DotNetBar.LabelX
    Friend WithEvents Txt_Nombre_Archivo As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents Circular_Progres_Porc As DevComponents.DotNetBar.Controls.CircularProgress
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Circular_Progres_Val As DevComponents.DotNetBar.Controls.CircularProgress
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents Btn_Buscar_Archivo As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Btn_Archivo_Ayuda_Excel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Btn_Cancelar As DevComponents.DotNetBar.ButtonItem
    Public WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Btn_Importar_Desde_XML As DevComponents.DotNetBar.ButtonItem
End Class

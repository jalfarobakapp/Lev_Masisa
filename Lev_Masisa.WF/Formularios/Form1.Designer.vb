<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Txt_Host = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Txt_Puerto = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Txt_Usuario = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Txt_Clave = New System.Windows.Forms.TextBox()
        Me.Btn_Lev_Participantes = New System.Windows.Forms.Button()
        Me.Lev_Transacciones = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Txt_Basededatos = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Txt_Host
        '
        Me.Txt_Host.Location = New System.Drawing.Point(80, 24)
        Me.Txt_Host.Name = "Txt_Host"
        Me.Txt_Host.Size = New System.Drawing.Size(194, 20)
        Me.Txt_Host.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(35, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Host"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(35, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Host"
        '
        'Txt_Puerto
        '
        Me.Txt_Puerto.Location = New System.Drawing.Point(80, 50)
        Me.Txt_Puerto.Name = "Txt_Puerto"
        Me.Txt_Puerto.Size = New System.Drawing.Size(194, 20)
        Me.Txt_Puerto.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(35, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Host"
        '
        'Txt_Usuario
        '
        Me.Txt_Usuario.Location = New System.Drawing.Point(80, 76)
        Me.Txt_Usuario.Name = "Txt_Usuario"
        Me.Txt_Usuario.Size = New System.Drawing.Size(194, 20)
        Me.Txt_Usuario.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(35, 105)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Host"
        '
        'Txt_Clave
        '
        Me.Txt_Clave.Location = New System.Drawing.Point(80, 102)
        Me.Txt_Clave.Name = "Txt_Clave"
        Me.Txt_Clave.Size = New System.Drawing.Size(194, 20)
        Me.Txt_Clave.TabIndex = 6
        '
        'Btn_Lev_Participantes
        '
        Me.Btn_Lev_Participantes.Location = New System.Drawing.Point(38, 175)
        Me.Btn_Lev_Participantes.Name = "Btn_Lev_Participantes"
        Me.Btn_Lev_Participantes.Size = New System.Drawing.Size(236, 23)
        Me.Btn_Lev_Participantes.TabIndex = 8
        Me.Btn_Lev_Participantes.Text = "Levantar PARTICIPANTES"
        Me.Btn_Lev_Participantes.UseVisualStyleBackColor = True
        '
        'Lev_Transacciones
        '
        Me.Lev_Transacciones.Location = New System.Drawing.Point(38, 204)
        Me.Lev_Transacciones.Name = "Lev_Transacciones"
        Me.Lev_Transacciones.Size = New System.Drawing.Size(236, 23)
        Me.Lev_Transacciones.TabIndex = 9
        Me.Lev_Transacciones.Text = "Levantar TRANSACCIONES"
        Me.Lev_Transacciones.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(35, 131)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(29, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Host"
        '
        'Txt_Basededatos
        '
        Me.Txt_Basededatos.Location = New System.Drawing.Point(80, 128)
        Me.Txt_Basededatos.Name = "Txt_Basededatos"
        Me.Txt_Basededatos.Size = New System.Drawing.Size(194, 20)
        Me.Txt_Basededatos.TabIndex = 10
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 258)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Txt_Basededatos)
        Me.Controls.Add(Me.Lev_Transacciones)
        Me.Controls.Add(Me.Btn_Lev_Participantes)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Txt_Clave)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Txt_Usuario)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Txt_Puerto)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Txt_Host)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Txt_Host As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Txt_Puerto As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Txt_Usuario As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Txt_Clave As TextBox
    Friend WithEvents Btn_Lev_Participantes As Button
    Friend WithEvents Lev_Transacciones As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Txt_Basededatos As TextBox
End Class

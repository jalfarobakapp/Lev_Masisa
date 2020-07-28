Public Class Form1
    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Cadena_ConexionSQL_Server = "data source = JALFARO-MSI\SQLEXPRESS; initial catalog = Datos_Masisa; user id = sa; password = antonia12.,"

    End Sub
    Private Sub Btn_Lev_Participantes_Click(sender As Object, e As EventArgs) Handles Btn_Lev_Participantes.Click


    End Sub

    Private Sub Lev_Transacciones_Click(sender As Object, e As EventArgs) Handles Lev_Transacciones.Click

        Dim Fm As New Frm_Levantar_Tablas
        Fm.ShowDialog(Me)
        Fm.Dispose()

    End Sub

End Class

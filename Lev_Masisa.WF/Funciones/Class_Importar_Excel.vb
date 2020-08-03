Imports Docs.Excel
Imports DevComponents.DotNetBar.Controls

Public Class Class_Importar_Excel

    Public Function Importar_Excel_Array(_Direccion_Archivo As String,
                                         _Extencion As String,
                                         Optional ByVal Hoja As Integer = 0,
                                         Optional ByRef _Barra_Progreso As Object = Nothing)

        ExcelWorkbook.SetLicenseCode("SA014N-E4113A-E1ALDA-101800")
        Dim Workbook As Object

        Dim Ext_ As String = LCase(_Extencion)

        Select Case Ext_
            Case "xls"
                Workbook = ExcelWorkbook.ReadXLS(_Direccion_Archivo)
            Case "xlsx"
                Workbook = ExcelWorkbook.ReadXLSX(_Direccion_Archivo)
            Case "csv"
                Workbook = ExcelWorkbook.ReadCSV(_Direccion_Archivo)
        End Select

        Dim Filas As Double = Workbook.Worksheets(Hoja).Rows.Count
        Dim Columnas As Double = Workbook.Worksheets(Hoja).Columns.Count

        Dim Arreglo(Filas - 1, Columnas - 1) As String

        If Not IsNothing(_Barra_Progreso) Then
            _Barra_Progreso.Maximum = Filas
        End If

        For i As Integer = 1 To Filas  ' Workbook.Worksheets(0).Rows.Count

            For cl As Integer = 0 To Columnas - 1
                Arreglo(i - 1, cl) = Workbook.Worksheets(Hoja).Cells(i - 1, cl).Value
            Next

            If Not IsNothing(_Barra_Progreso) Then
                _Barra_Progreso.Value += 1
                System.Windows.Forms.Application.DoEvents()
            End If

        Next i

        Return Arreglo

    End Function

End Class

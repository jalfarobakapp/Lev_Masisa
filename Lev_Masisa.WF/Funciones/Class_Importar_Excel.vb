Imports Docs.Excel

Public Class Class_Importar_Excel

    Public Function Importar_Excel_Array(ByVal Direccion_Archivo_XLS As String,
                                         ByVal Extencion_ As String,
                                         Optional ByVal Hoja As Integer = 0)

        ExcelWorkbook.SetLicenseCode("SA014N-E4113A-E1ALDA-101800")
        Dim Workbook As Object

        Dim Ext_ As String = LCase(Extencion_)

        If Ext_ = "xls" Then
            Workbook = ExcelWorkbook.ReadXLSX(Direccion_Archivo_XLS)
        ElseIf Ext_ = "xlsx" Then
            Workbook = ExcelWorkbook.ReadXLSX(Direccion_Archivo_XLS)
        End If

        Dim Filas As Double = Workbook.Worksheets(Hoja).Rows.Count
        Dim Columnas As Double = Workbook.Worksheets(Hoja).Columns.Count

        Dim Arreglo(Filas - 1, Columnas - 1) As String


        Dim dt As New DataTable
        dt.Columns.Add("Codigo")
        dt.Columns.Add("Precio")

        For i As Integer = 1 To Filas  ' Workbook.Worksheets(0).Rows.Count

            For cl As Integer = 0 To Columnas - 1
                Arreglo(i - 1, cl) = Workbook.Worksheets(Hoja).Cells(i - 1, cl).Value
            Next

        Next i

        Return Arreglo

    End Function

End Class

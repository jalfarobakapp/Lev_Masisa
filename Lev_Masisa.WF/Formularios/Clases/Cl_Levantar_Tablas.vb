Imports DevComponents.DotNetBar
Imports DevComponents.DotNetBar.Controls

Public Class Cl_Levantar_Tablas

    Dim _Sql As New Class_SQL(Cadena_ConexionSQL_Server)
    Dim Consulta_Sql As String

    Public Sub New()

    End Sub

    Function Sb_Importar_Archivo_Excel(_Formulario As Form) As String(,) '--List(Of String)

        Dim _Nombre_Archivo As String
        Dim _Ubic_Archivo As String

        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            '.Filter = "Ficheros DBF (PFMDSP10.dbf)|PFMDSP10.dbf|Todos (*.*)|*.*"
            .Filter = "Ficheros (*.csv)|*.csv|Todos los archivos (*.*)|*.*"
            'Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
            .FileName = String.Empty
            '.ShowDialog(Me)

            If .ShowDialog(_Formulario) = DialogResult.OK Then

                _Nombre_Archivo = System.IO.Path.GetFileNameWithoutExtension(.SafeFileName)
                _Ubic_Archivo = System.IO.Path.GetDirectoryName(.FileName)

                _Nombre_Archivo = .SafeFileName
                _Ubic_Archivo = .FileName
            Else
                Return Nothing
            End If

        End With

        Dim _ImpEx As New Class_Importar_Excel
        Dim _Extencion As String = Replace(System.IO.Path.GetExtension(_Nombre_Archivo), ".", "")
        Dim _Arreglo = _ImpEx.Importar_Excel_Array(_Ubic_Archivo, _Extencion, 0)
        'Dim _Filas = _Arreglo.GetUpperBound(0)

        'Dim _Arreglo_Cd(_Filas) As String

        Return _Arreglo

    End Function

    Function Fx_Importar_Tabla_Transacciones(_Arreglo As Object,
                                             _Primera_Fila_Es_Encabezado As Boolean,
                                             ByRef _Leyenda As String,
                                             ByRef _Circular_Progres_Porc As CircularProgress,
                                             ByRef _Circular_Progres_Val As CircularProgress) As Boolean

        Dim _Desde = 0

        If _Primera_Fila_Es_Encabezado Then
            _Desde = 1
        End If

        'For i = _Desde To _Filas
        '    _Arreglo_Cd(i) = NuloPorNro(_Arreglo(i, 0), "")
        'Next

        Dim _Problemas As Integer
        Dim _SinProbremas As Integer

        'Sb_Habilitar_Deshabilitar_Comandos(False, True)
        Dim _Contador As Integer = 0

        Dim _SqlQuery As String = String.Empty
        Dim _SqlLotes As New List(Of String)
        Dim _Lotes = 1000

        Dim _Filas = _Arreglo.GetUpperBound(0)

        _Circular_Progres_Val.Maximum = _Filas

        For i = _Desde To _Filas

            Dim _Error = String.Empty

            System.Windows.Forms.Application.DoEvents()

#Region "VARIABLES"


            Dim _Id_Transacciones As String
            Dim _Id As String
            Dim _Carga_ID As String
            Dim _Archivo As String
            Dim _Pais As String
            Dim _PLC_ID As String
            Dim _Suc_ID As String
            Dim _Beneficiario_ID As String
            Dim _PLC_Nombre As String
            Dim _Suc_Nombre As String
            Dim _Codigo_PLC As String
            Dim _Codigo_Suc As String
            Dim _Tipo As String
            Dim _Numero As String
            Dim _Fecha As String
            Dim _Grabacion As String
            Dim _Vendedor As String
            Dim _Nombre_Vendedor As String
            Dim _Cliente As String
            Dim _Nombre_Cliente As String
            Dim _Mueblista As String
            Dim _Nombre_Mueblista As String
            Dim _Codigo As String
            Dim _Codigo_Masisa As String
            Dim _Descripcion_Producto As String
            Dim _Ud As String
            Dim _Cant As Double
            Dim _P_Bruto_Lista As Double
            Dim _P_Bruto_Venta As Double
            Dim _V_Bruto_Venta As Double
            Dim _Doc_Origen As String
            Dim _Marca As String
            Dim _Sup_Fam As String
            Dim _Familia As String
            Dim _Sub_Fam As String
            Dim _Cat_Masisa As String
            Dim _Cat_Masisa_2 As String
            Dim _Puntos As Double
            Dim _Cartola As String
            Dim _Fecha_Procesado As Date
            Dim _Procesado As Boolean
            Dim _Estado As Double
            Dim _Regla As String
            Dim _Fecha_Creacion As Date
            Dim _Rut_Final As String
            Dim _Periodo_Activo As Boolean
            Dim _Red_M As Boolean

#End Region

            Try

                _Id_Transacciones = NuloPorNro(_Arreglo(i, 0), "")
                _Id = NuloPorNro(_Arreglo(i, 1), "")
                _Carga_ID = NuloPorNro(_Arreglo(i, 2), "")
                _Archivo = NuloPorNro(_Arreglo(i, 3), "")
                _Pais = NuloPorNro(_Arreglo(i, 4), "")
                _PLC_ID = NuloPorNro(_Arreglo(i, 5), "")
                _Suc_ID = NuloPorNro(_Arreglo(i, 6), "")
                _Beneficiario_ID = NuloPorNro(_Arreglo(i, 7), "")
                _PLC_Nombre = NuloPorNro(_Arreglo(i, 8), "")
                _Suc_Nombre = NuloPorNro(_Arreglo(i, 9), "")
                _Codigo_PLC = NuloPorNro(_Arreglo(i, 10), "")
                _Codigo_Suc = NuloPorNro(_Arreglo(i, 11), "")
                _Tipo = NuloPorNro(_Arreglo(i, 12), "")
                _Numero = NuloPorNro(_Arreglo(i, 13), "")

                _Fecha = NuloPorNro(_Arreglo(i, 14), #1/1/2000#)
                _Grabacion = NuloPorNro(_Arreglo(i, 15), #1/1/2000#)

                _Vendedor = NuloPorNro(_Arreglo(i, 16), "")
                _Nombre_Vendedor = NuloPorNro(_Arreglo(i, 17), "")
                _Cliente = NuloPorNro(_Arreglo(i, 18), "")
                _Nombre_Cliente = NuloPorNro(_Arreglo(i, 19), "")
                _Mueblista = NuloPorNro(_Arreglo(i, 20), "")
                _Nombre_Mueblista = NuloPorNro(_Arreglo(i, 21), "")
                _Codigo = NuloPorNro(_Arreglo(i, 22), "")
                _Codigo_Masisa = NuloPorNro(_Arreglo(i, 23), "")
                _Descripcion_Producto = NuloPorNro(_Arreglo(i, 24), "")
                _Ud = NuloPorNro(_Arreglo(i, 25), "")

                _Cant = NuloPorNro(_Arreglo(i, 26), 0)
                _P_Bruto_Lista = NuloPorNro(_Arreglo(i, 27), 0)
                _P_Bruto_Venta = NuloPorNro(_Arreglo(i, 28), 0)
                _V_Bruto_Venta = NuloPorNro(_Arreglo(i, 29), 0)

                _Doc_Origen = NuloPorNro(_Arreglo(i, 30), "")
                _Marca = NuloPorNro(_Arreglo(i, 31), "")
                _Sup_Fam = NuloPorNro(_Arreglo(i, 32), "")
                _Familia = NuloPorNro(_Arreglo(i, 33), "")
                _Sub_Fam = NuloPorNro(_Arreglo(i, 34), "")

                _Cat_Masisa = NuloPorNro(_Arreglo(i, 35), "")
                _Cat_Masisa_2 = NuloPorNro(_Arreglo(i, 36), "")
                _Puntos = NuloPorNro(_Arreglo(i, 37), 0)

                _Cartola = NuloPorNro(_Arreglo(i, 38), "")
                _Fecha_Procesado = NuloPorNro(_Arreglo(i, 39), #1/1/2000#)

                _Procesado = NuloPorNro(_Arreglo(i, 40), 0)
                _Estado = NuloPorNro(_Arreglo(i, 41), 0)
                _Regla = NuloPorNro(_Arreglo(i, 42), "")

                _Fecha_Creacion = NuloPorNro(_Arreglo(i, 43), #1/1/2000#)

                _Rut_Final = NuloPorNro(_Arreglo(i, 44), "")
                _Periodo_Activo = NuloPorNro(_Arreglo(i, 45), 0)
                _Red_M = NuloPorNro(_Arreglo(i, 46), 0)

                '_Fecha_Docto = NuloPorNro(_Arreglo(i, 6), #1/1/2000#)

            Catch ex As Exception
                _Error = ex.Message
            End Try

            If String.IsNullOrEmpty(_Error) Then

                _SqlQuery += "Insert Into Transacciones (Id, [Carga ID], Archivo, Pais, [PLC ID], [Suc ID], [Beneficiario ID], [PLC Nombre], [Suc Nombre], " &
                             "[Codigo PLC], [Codigo Suc], Tipo, Numero, Fecha, Grabacion, Vendedor, [Nombre Vendedor], Cliente, [Nombre Cliente], " &
                             "Mueblista, [Nombre Mueblista], Codigo, [Codigo Masisa], [Descripcion Producto], Ud, Cant, [P Bruto Lista], [P Bruto Venta], " &
                             "[V Bruto Venta], [Doc Origen], Marca, [Sup Fam], Familia, [Sub Fam], [Cat Masisa], [Cat Masisa 2], Puntos, Cartola, " &
                             "[Fecha Procesado], Procesado, Estado, Regla, [Fecha Creacion], [Rut Final], [Periodo Activo], [Red M]) Values " &
                             "('" & _Id & "','" & _Carga_ID & "','" & _Archivo & "','" & _Pais & "','" & _PLC_ID &
                             "','" & _Suc_ID & "','" & _Beneficiario_ID & "','" & _PLC_Nombre & "','" & _Suc_Nombre & "', " &
                             "'" & _Codigo_PLC & "','" & _Codigo_PLC & "','" & _Tipo & "','" & _Numero &
                             "','" & Format(_Fecha, "yyyyMMdd") & "','" & Format(_Grabacion, "yyyyMMdd") &
                             "','" & _Vendedor & "','" & _Nombre_Vendedor & "','" & _Cliente & "','" & _Nombre_Cliente & "', " &
                             "'" & _Mueblista & "','" & _Nombre_Mueblista & "','" & _Codigo & "','" & _Codigo_Masisa & "','" & _Descripcion_Producto &
                             "','" & _Ud & "'," & De_Num_a_Tx_01(_Cant, False, 5) &
                             "," & De_Num_a_Tx_01(_P_Bruto_Lista, False, 5) &
                             "," & De_Num_a_Tx_01(_P_Bruto_Lista, False, 5) &
                             ", " & De_Num_a_Tx_01(_V_Bruto_Venta, False, 5) &
                             ",'" & _Doc_Origen & "','" & _Marca & "','" & _Sup_Fam & "','" & _Familia & "','" & _Sub_Fam &
                             "','" & _Cat_Masisa & "','" & _Cat_Masisa_2 &
                             "'," & De_Num_a_Tx_01(_Puntos, False, 5) &
                             ",'" & _Cartola & "','" & Format(_Fecha_Procesado, "yyyyMMdd") & "'," & Convert.ToInt32(_Procesado) & "," & _Estado &
                             ",'" & _Regla & "','" & Format(_Fecha_Creacion, "yyyyMMdd") &
                             "','" & _Rut_Final & "'," & Convert.ToInt32(_Periodo_Activo) & "," & Convert.ToInt32(_Red_M) & ")" & vbCrLf

            Else
                'Sb_AddToLog("Fila Nro :" & i + 1, "Problema: " & _Error & "Código: [" & _Kopr & "]", _
                ' _Txt_Log, False)
                _Problemas += 1
            End If

            If _Contador >= _Lotes Then
                If _Contador Mod _Lotes = 0 Then
                    _SqlLotes.Add(_SqlQuery)
                    _SqlQuery = String.Empty
                End If
            End If

            If CBool(_Problemas) Then
                _Circular_Progres_Porc.ProgressColor = Color.Red
                _Circular_Progres_Val.ProgressColor = Color.Red
            End If

            'If _Cancelar Then
            '    Exit For
            'End If

            System.Windows.Forms.Application.DoEvents()

            _Contador += 1
            _Circular_Progres_Porc.Value = ((_Contador * 100) / _Filas)
            _Circular_Progres_Val.Value += 1
            _Circular_Progres_Val.ProgressText = _Circular_Progres_Val.Value

            _Leyenda = "Leyendo fila " & i & " de " & _Filas & ". Estado Ok: " & _SinProbremas & ", Problemas: " & _Problemas

        Next

        If Not String.IsNullOrEmpty(_SqlQuery) Then
            _SqlLotes.Add(_SqlQuery)
        End If

        '_SqlQuery = "Delete " & _Global_BaseBk & "Zw_Compras_en_SII Where Periodo = " & _Periodo & " And Mes = " & _Mes &
        '            vbCrLf &
        '            vbCrLf &
        '            _SqlQuery

        For Each _SqlQl As String In _SqlLotes

            _Sql = New Class_SQL(Cadena_ConexionSQL_Server)

            'Dim _Cn As New SqlClient.SqlConnection
            '_Sql.Sb_Abrir_Conexion(_Cn)

            _Sql.Fx_Eje_Condulta_Insert_Update_Delte_TRANSACCION(_SqlQl)

        Next



        'Consulta_Sql = "Update " & _Global_BaseBk & "Zw_Compras_en_SII Set" & vbCrLf &
        '               "Monto_Exento = Monto_Exento*-1," & vbCrLf &
        '               "Monto_Neto = Monto_Neto*-1," & vbCrLf &
        '               "Monto_Iva_Recuperable = Monto_Iva_Recuperable *-1," & vbCrLf &
        '               "Monto_Iva_No_Recuperable = Monto_Iva_No_Recuperable *-1," & vbCrLf &
        '               "Monto_Total = Monto_Total*-1," & vbCrLf &
        '               "Valor_Otro_impuesto = Valor_Otro_impuesto*-1," & vbCrLf &
        '               "Vanedo = Vanedo*-1," & vbCrLf &
        '               "Vaivdo = Vaivdo*-1," & vbCrLf &
        '               "Vabrdo = Vabrdo*-1" & vbCrLf &
        '               "Where Periodo = " & _Periodo & " And Mes = " & _Mes & " And Tido = 'NCC'"
        '_Sql.Ej_consulta_IDU(Consulta_Sql)

    End Function

End Class

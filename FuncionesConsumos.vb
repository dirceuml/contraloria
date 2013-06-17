Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesConsumos

    Private Function LlenaConsumosInactivos(ByVal IdPeriodo As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet
        Dim IdAño, IdMes As Integer
        Dim IdPeriodoAnt As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
        If IdMes > 1 Then IdPeriodoAnt = (IdAño * IdMes - 1).ToString Else IdPeriodoAnt = ((IdAño - 1) * 100 + 12).ToString

        sql = "select IdPeriodo, NroContrato, Cliente, MontoNetoContratoUSD, MontoConsumidoUSD, FactLetrasEmitidasUSD, MontoCobradoUSD, FormaPago, FlagInactivo " & _
                "from Consumo where FlagInactivo = 'S' and IdPeriodo = " & IdPeriodo
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        If dtaset.Tables(0).Rows.Count = 0 Then
            sql = "select " & IdPeriodo & " as IdPeriodo, NroContrato, Cliente, MontoNetoContratoUSD, MontoConsumidoUSD, FactLetrasEmitidasUSD, MontoCobradoUSD, FormaPago, FlagInactivo " & _
                "from Consumo where FlagInactivo = 'S' and IdPeriodo = " & IdPeriodoAnt
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
        End If

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub GeneraPlantillaConsumosInactivos(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet

        Dim array As Object(,)
        Dim dt As DataTable
        Dim cant As Integer

        'Try
        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("Consumos")

        dt = LlenaConsumosInactivos(IdPeriodo)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, False)
        hoja.Range("A2:I" & (cant + 1).ToString, Type.Missing).Value2 = array

        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(reporte)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        'hoja = Nothing
        reporte = Nothing
        excel = Nothing
        GC.Collect()
        'Dim proc As System.Diagnostics.Process
        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub EliminaInactivos(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from Consumo where FlagInactivo = 'S' and IdPeriodo = " & IdPeriodo
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ProcesoCargaConsumos(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_ProcesoCargaConsumos"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdPeriodo", IdPeriodo)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraReportesConsumos(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        GeneraResumen(reporte, IdPeriodo)
        GeneraCobradosPorConsumir(reporte, IdPeriodo, 1)
        GeneraConsumosPorCobrar(reporte, IdPeriodo, 1)

        reporte.Worksheets("Plantilla").Delete()
        reporte.Worksheets("Plantilla2").Delete()
        reporte.Worksheets("Plantilla3").Delete()

        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(reporte)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        'hoja = Nothing
        reporte = Nothing
        excel = Nothing
        GC.Collect()
        'Dim proc As System.Diagnostics.Process
        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub GeneraResumen(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim hoja As Excel.Worksheet

        Dim array As Object(,)

        Dim fil_ini, Cant As Integer

        hoja = reporte.Worksheets("RESUMEN")
        hoja.Cells(5, 2).Value = "A " & BuscaPeriodo(IdPeriodo).ToUpper

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        fil_ini = 10
        sql = "select 'AÑO ' + convert(varchar(4), IdAñoContrato), sum(SaldoSobrePagadoUSD) as SaldoSobrePagadoUSD, sum(SaldoSobrePagadoUSD) - sum(SaldoCobradoPorConsumirUSD) as SaldoPorCobrarUSD, " & _
                "sum(SaldoCobradoPorConsumirUSD) as SaldoCobradoPorConsumirUSD from Consumo " & _
                "where IdPeriodo = " & IdPeriodo & " group by IdAñoContrato "
        cmd = New SqlCommand(sql, cn)
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count
        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        If fil_ini + Cant <= 19 Then hoja.Rows((fil_ini + Cant).ToString & ":19").EntireRow.Hidden = True

        fil_ini = 24
        sql = "select 'AÑO ' + convert(varchar(4), IdAñoContrato), isnull(sum(SaldoFactLetrasUSD), 0) as SaldoFactLetrasUSD, 0 as SaldoPorCobrarUSD, " & _
                "isnull(sum(SaldoFactLetrasUSD), 0) as SaldoFactLetrasUSD from Consumo " & _
                "where IdPeriodo = " & IdPeriodo & " group by IdAñoContrato "
        cmd = New SqlCommand(sql, cn)
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count
        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        If fil_ini + Cant <= 33 Then hoja.Rows((fil_ini + Cant).ToString & ":33").EntireRow.Hidden = True

        cn.Close()
    End Sub

    Sub GeneraCobradosPorConsumir(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim rdr As SqlDataReader
        Dim dt As DataTable
        Dim sql As String

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim array As Object(,)

        Dim aux As String

        Dim IdAñoContrato As Integer
        Dim fil_ini2, i, i2, Cant As Integer

        hoja = reporte.Worksheets("COBRADOS POR CONSUMIR")
        hoja.Cells(4, 2).Value = "CONTRATOS FACTURADOS POR CONSUMIR A " & BuscaPeriodo(IdPeriodo).ToUpper

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(C.IdAñoContrato, T.IdAñoContrato) as IdAñoContrato, " & _
                "isnull(C.FlagInactivo, T.FlagInactivo) as FlagInactivo, isnull(Cant,0) as Cant from " & _
                "(select IdAñoContrato, isnull(FlagInactivo, 'N') as FlagInactivo, count(*) as Cant from Consumo where IdPeriodo = " & IdPeriodo & " " & _
                "group by IdAñoContrato, isnull(FlagInactivo, 'N')) C full join " & _
                "(select distinct IdAñoContrato, T.FlagInactivo from Consumo, (select 'N' as FlagInactivo union select 'S') T where IdPeriodo = " & IdPeriodo & ") T " & _
                "on C.IdAñoContrato = T.IdAñoContrato and C.FlagInactivo = T.FlagInactivo " & _
                "order by 1, 2"
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        While rdr.Read
            IdAñoContrato = Convert.ToInt32(rdr("IdAñoContrato"))
            Cant = Convert.ToInt32(rdr("Cant"))

            sql = "select NroContrato, Cliente, MontoNetoContratoUSD, MontoConsumidoUSD, FactLetrasEmitidasUSD, isnull(SaldoSobrePagadoUSD, 0) as SaldoSobrePagadoUSD, FormaPago " &
                    "from Consumo C where IdPeriodo = " & IdPeriodo & " and IdAñoContrato = " & IdAñoContrato.ToString & " " & _
                    "and isnull(FlagInactivo, 'N') = '" & rdr("FlagInactivo") & "' " & _
                    "order by NroContrato"
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)

            If rdr("FlagInactivo") = "N" Then
                aux = "ACTIVOS"
                i = 0
                If Cant > 0 Then 'Con Clientes Activos
                    While i < Cant
                        If IdAñoContrato >= 2011 Then
                            rango = hoja.Rows("1:6")
                            rango.Copy(hoja.Rows(fil_ini))
                            'If IdAñoContrato = 2009 Then hoja.Rows(fil_ini.ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True
                            hoja.Range("B" & (fil_ini + 3).ToString & ":H" & (fil_ini + 3).ToString).MergeCells = True
                            rango = reporte.Worksheets("Plantilla").Rows("2:65")
                            rango.Copy(hoja.Rows(fil_ini + 6))
                            If i = 0 Then
                                hoja.Cells(fil_ini + 7, 2).Value = "PERIODO " & IdAñoContrato.ToString & " - " & aux
                            Else
                                hoja.Cells(fil_ini + 7, 2).Value = "VIENEN.."
                                hoja.Cells(fil_ini + 7, 4).FormulaR1C1 = "=R[-9]C[0]"
                                hoja.Cells(fil_ini + 7, 5).FormulaR1C1 = "=R[-9]C[0]"
                                hoja.Cells(fil_ini + 7, 6).FormulaR1C1 = "=R[-9]C[0]"
                                hoja.Cells(fil_ini + 7, 7).FormulaR1C1 = "=R[-9]C[0]"
                            End If
                        Else
                            rango = hoja.Rows("1:6")
                            rango.Copy(hoja.Rows(fil_ini))
                            If IdAñoContrato = 2007 Or IdAñoContrato = 2008 Then hoja.Rows(fil_ini.ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True
                            hoja.Range("B" & (fil_ini + 3).ToString & ":H" & (fil_ini + 3).ToString).MergeCells = True
                            rango = reporte.Worksheets("Plantilla2").Rows("2:79")
                            rango.Copy(hoja.Rows(fil_ini + 6))
                            hoja.Cells(fil_ini + 7, 2).Value = "PERIODO " & IdAñoContrato.ToString & " - " & aux
                        End If

                        fil_ini2 = fil_ini + 70
                        fil_ini = fil_ini + 8

                        If i + 60 > Cant - 1 Then i2 = Cant - 1 Else i2 = i + 60 - 1
                        array = DataSet2Array(dt, i, i2)
                        hoja.Range("B" & fil_ini.ToString & ":H" & (fil_ini + (i2 - i)).ToString, Type.Missing).Value2 = array

                        If i2 = Cant - 1 Then
                            If IdAñoContrato >= 2011 Then
                                hoja.Rows((fil_ini + i2 - i + 1).ToString & ":" & (fil_ini + 60 - 1).ToString).EntireRow.Hidden = True
                                hoja.Cells(fil_ini + 60, 3).Value = "TOTAL " & IdAñoContrato.ToString & " US$"
                            Else
                                hoja.Rows((fil_ini + i2 - i + 1).ToString & ":" & (fil_ini + 60 - 1).ToString).EntireRow.Hidden = True
                            End If
                        Else
                            If IdAñoContrato >= 2011 Then
                                hoja.Cells(fil_ini + 60, 3).Value = "VAN..."
                            End If
                        End If
                        If IdAñoContrato = 2009 Or IdAñoContrato = 2010 Then
                            hoja.HPageBreaks.Add(hoja.Rows(fil_ini + 76))
                        ElseIf IdAñoContrato >= 2011 Then
                            hoja.HPageBreaks.Add(hoja.Rows(fil_ini + 62))
                        End If

                        i = i + 60
                        If IdAñoContrato >= 2011 Then
                            fil_ini = fil_ini + 62
                        Else
                            fil_ini = fil_ini + 76
                        End If

                    End While
                Else
                    'Sin Clientes Activos
                    rango = hoja.Rows("1:6")
                    rango.Copy(hoja.Rows(fil_ini))
                    hoja.Range("B" & (fil_ini + 3).ToString & ":H" & (fil_ini + 3).ToString).MergeCells = True
                    rango = reporte.Worksheets("Plantilla2").Rows("2:79")
                    rango.Copy(hoja.Rows(fil_ini + 6))
                    hoja.Rows((fil_ini + 7).ToString & ":" & (fil_ini + 68).ToString).EntireRow.Hidden = True
                    fil_ini2 = fil_ini + 70
                    fil_ini = fil_ini + 94
                End If
            Else
                aux = "INACTIVOS"
                If Cant > 0 Then
                    hoja.Cells(fil_ini2 - 1, 2).Value = "PERIODO " & IdAñoContrato.ToString & " - " & aux
                    array = DataSet2Array(dt, False)
                    hoja.Range("B" & fil_ini2.ToString & ":H" & (fil_ini2 + Cant - 1).ToString, Type.Missing).Value2 = array
                    hoja.Rows((fil_ini2 + Cant).ToString & ":" & (fil_ini2 + 10 - 1).ToString).EntireRow.Hidden = True
                    hoja.Cells(fil_ini2 + 12, 3).Value = "TOTAL " & IdAñoContrato.ToString & " US$"
                Else
                    hoja.Cells(fil_ini2 + 12, 3).Value = "TOTAL " & IdAñoContrato.ToString & " US$"
                    hoja.Rows((fil_ini2 - 1).ToString & ":" & (fil_ini2 + 10).ToString).EntireRow.Hidden = True
                End If
                If IdAñoContrato = 2008 Then hoja.HPageBreaks.Add(hoja.Rows(fil_ini))
            End If
        End While

        hoja.PageSetup.PrintArea = "$B$1:$H$" + CStr(fil_ini - 1)

        cn.Close()
    End Sub

    Sub GeneraConsumosPorCobrar(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim rdr As SqlDataReader
        Dim dt As DataTable
        Dim sql As String

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim array As Object(,)

        Dim IdAñoContrato As Integer
        Dim fil_ini2, i, i2, Cant As Integer

        hoja = reporte.Worksheets("CONSUMOS POR COBRAR")
        hoja.Cells(4, 2).Value = "CONSUMOS POR FACTURAR A " & BuscaPeriodo(IdPeriodo).ToUpper

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdAñoContrato, count(*) as Cant from Consumo " & _
                "where IdPeriodo = " & IdPeriodo & " and SaldoFactLetrasUSD > 0" & _
                "group by IdAñoContrato order by 1"
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        While rdr.Read
            IdAñoContrato = Convert.ToInt32(rdr("IdAñoContrato"))
            Cant = Convert.ToInt32(rdr("Cant"))

            sql = "select NroContrato, Cliente, MontoNetoContratoUSD, MontoConsumidoUSD, FactLetrasEmitidasUSD, isnull(SaldoFactLetrasUSD, 0) as SaldoFactLetrasUSD " &
                    "from Consumo C where IdPeriodo = " & IdPeriodo & " and IdAñoContrato = " & IdAñoContrato.ToString & " and SaldoFactLetrasUSD > 0 " & _
                    "order by NroContrato"
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)

            i = 0
            While i < Cant
                rango = hoja.Rows("1:6")
                rango.Copy(hoja.Rows(fil_ini))
                hoja.Range("B" & (fil_ini + 3).ToString & ":H" & (fil_ini + 3).ToString).MergeCells = True
                rango = reporte.Worksheets("Plantilla3").Rows("2:55")
                rango.Copy(hoja.Rows(fil_ini + 6))
                If i = 0 Then
                    hoja.Cells(fil_ini + 7, 2).Value = "PERIODO " & IdAñoContrato.ToString
                Else
                    hoja.Cells(fil_ini + 7, 2).Value = "VIENEN.."
                    hoja.Cells(fil_ini + 7, 4).FormulaR1C1 = "=R[-9]C[0]"
                    hoja.Cells(fil_ini + 7, 5).FormulaR1C1 = "=R[-9]C[0]"
                    hoja.Cells(fil_ini + 7, 6).FormulaR1C1 = "=R[-9]C[0]"
                    hoja.Cells(fil_ini + 7, 7).FormulaR1C1 = "=R[-9]C[0]"
                End If
                If IdAñoContrato >= 2007 And IdAñoContrato <= 2010 Then hoja.Rows(fil_ini.ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True

                fil_ini2 = fil_ini + 30
                fil_ini = fil_ini + 8

                If i + 50 > Cant - 1 Then i2 = Cant - 1 Else i2 = i + 50 - 1
                array = DataSet2Array(dt, i, i2)
                hoja.Range("B" & fil_ini.ToString & ":G" & (fil_ini + (i2 - i)).ToString, Type.Missing).Value2 = array

                If i2 = Cant - 1 Then
                    hoja.Rows((fil_ini + i2 - i + 1).ToString & ":" & (fil_ini + 50 - 1).ToString).EntireRow.Hidden = True
                    hoja.Cells(fil_ini + 50, 3).Value = "TOTAL " & IdAñoContrato.ToString & " US$"
                Else
                    hoja.Cells(fil_ini + 50, 3).Value = "VAN..."
                End If
                i = i + 50
                fil_ini = fil_ini + 52
                If IdAñoContrato >= 2010 Then hoja.HPageBreaks.Add(hoja.Rows(fil_ini))
            End While
        End While

        hoja.PageSetup.PrintArea = "$B$1:$H$" + CStr(fil_ini - 1)

        cn.Close()
    End Sub

End Module

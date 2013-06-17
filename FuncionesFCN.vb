Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Configuration

Module FuncionesFCN
    Dim excel As Excel.Application

    Sub GeneraReportesFCN(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        excel = New Excel.Application
        excel.DisplayAlerts = False

        FuncionesRep.LlenaITF()

        reporte = excel.Workbooks.Open(RutaPlantilla)
        hoja = reporte.Worksheets("FLUJO")
        CreaReporteFlujoCajaN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("INGRESOS")
        CreaReporteIngresosN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COMISION AG")
        CreaReporteAgenciasN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("BANCOS")
        CreaReporteBancosN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COSTOS I.")
        CreaReporteCostosIndirectosN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("SERVICIOS")
        CreaReporteServiciosN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("PROG. NACIONALES")
        CreaReporteProgramasNacionalesN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("SUELDOS Y HONO")
        CreaReporteSueldosN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("CARGAS SOCIALES")
        CreaReporteCargasSocialesN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("OTRAS CARGAS PERSONAL")
        CreaReporteOtrasCargasPersonalN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("MTC Y MUNICIP")
        CreaReporteMTCMunicipalidadesN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("RENTA E IGV")
        CreaReporteSUNATN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("GRUPO ATV")
        CreaReporteGrupoATVN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("INVERSIONES")
        CreaReporteInversionesN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("MATERIAL FILM")
        CreaReporteMaterialFilmicoN(hoja, IdPeriodo)
        hoja = reporte.Worksheets("MTTOS")
        CreaReporteMantenimientosN(hoja, IdPeriodo)
        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        excel = Nothing
    End Sub

    Sub CreaReporteFlujoCajaN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño As Integer

        Try
            hoja.Activate()
            hoja.Cells(4, 2).Value = ("Flujo de Caja a " + BuscaPeriodo(IdPeriodo)).ToUpper()

            IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
            'hoja.Cells(9, 3).Value = BuscaSaldoFinal(False, ((IdAño - 1) * 100 + 12).ToString)
            hoja.Cells(9, 3).Value = BuscaCierreCaja(Convert.ToDateTime((IdAño - 1).ToString & "-12-31"))

            GeneraResumenN(hoja, IdPeriodo, "INGRES", 12, 13, 2)
            GeneraResumenN(hoja, IdPeriodo, "DESING", 15, 16, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC01", 20, 24, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC02", 26, 33, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC03", 35, 40, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC04", 42, 43, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC05", 45, 50, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC06", 52, 71, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC07", 84, 84, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC08", 86, 110, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC09", 111, 115, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC10", 117, 122, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC11", 124, 127, 2)
            GeneraResumenN(hoja, IdPeriodo, "EGRC12", 129, 130, 2)

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteIngresosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "INGPUB", 9, 13, 2)
            GeneraDetalleN(hoja, IdPeriodo, "INGOTR", 21, 22, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteAgenciasN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "AGCNAC", 9, 57, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteBancosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalle2N(hoja, IdPeriodo, "BANCOS", 9, 9, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:D").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteCostosIndirectosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "COSIND", 9, 48, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteServiciosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "SERVIC", 9, 28, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteProgramasNacionalesN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "PRGNAC", 9, 28, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteSueldosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "SUELDO", 9, 18, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteCargasSocialesN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalle2N(hoja, IdPeriodo, "CARSOC", 9, 18, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:D").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteOtrasCargasPersonalN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "OTRCAR", 9, 16, 2)
            GeneraDetalleN(hoja, IdPeriodo, "BONOS", 24, 24, 2)
            GeneraDetalleN(hoja, IdPeriodo, "UTLANT", 33, 33, 2)
            GeneraDetalleN(hoja, IdPeriodo, "UTLACT", 34, 34, 2)
            GeneraDetalleN(hoja, IdPeriodo, "UTLFUT", 35, 35, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteMTCMunicipalidadesN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalle2N(hoja, IdPeriodo, "MTC", 9, 13, 2)
            GeneraDetalle2N(hoja, IdPeriodo, "MUNICP", 21, 25, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:D").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteSUNATN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalle2N(hoja, IdPeriodo, "SUNAT", 9, 20, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:D").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteGrupoATVN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleGrupoATVN(hoja, IdPeriodo, 10, 34, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteInversionesN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "INVERS", 9, 18, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteMaterialFilmicoN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(6, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalle2N(hoja, IdPeriodo, "MATFIL", 9, 48, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:D").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteMantenimientosN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            hoja.Cells(7, 2).Value = ("A " + BuscaPeriodo(IdPeriodo)).ToUpper()

            GeneraDetalleN(hoja, IdPeriodo, "MANT", 10, 14, 2)

            hoja.Columns("C:G").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub GeneraResumenN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            sql = "select Rubro, sum(Signo * MontoBruto) as BrutoAcum, -1 * sum(Signo * IGV) as IGVAcum " & _
                    "from V_FlujoN where CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " " &
                    "group by Orden, Orden2, Rubro order by Orden, Orden2, Rubro"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 2
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()

            If CodSeccion.Substring(0, 4) = "EGRC" Then hoja.Cells(fil_ini + i - 1, col_ini + 4).FormulaR1C1 = "=SUM(R" & fil_ini.ToString & "C[-1]:R" & (fil_ini + i - 1).ToString & "C[-1])"
            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraDetalleN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            sql = "select A.Rubro, BrutoAcum - isnull(MontoBruto, 0) as BrutoAnt, -1 * (IGVAcum - isnull(IGV, 0)) as IGVAnt, NetoAcum - isnull(MontoNeto, 0) as NetoAnt, -1 * isnull(IGV, 0) as IGV, isnull(MontoNeto, 0) as MontoNeto " & _
                    "from (select Orden2, Rubro, sum(Signo * MontoBruto) as BrutoAcum, sum(Signo * IGV) as IGVAcum, sum(Signo * MontoNeto) as NetoAcum " & _
                    "from V_DetalleN where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden2, Rubro) A " & _
                    "left join (select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV, sum(Signo * MontoBruto) as MontoBruto from V_DetalleN where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " group by Rubro) M on A.Rubro = M.Rubro " & _
                    "order by A.Orden2, A.Rubro"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                If (rdr("rubro").ToString <> "Sub Total US$") Then
                    For j = 0 To 5
                        hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                    Next
                End If
                i = i + 1
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraDetalle2N(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            sql = "select A.Rubro, BrutoAcum - isnull(MontoBruto, 0) as BrutoAnt, isnull(MontoBruto, 0) as MontoBruto " & _
                    "from (select Orden2, Rubro, sum(MontoBruto) as BrutoAcum " & _
                    "from V_DetalleN where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden2, Rubro) A " & _
                    "left join (select Rubro, sum(Signo * MontoBruto) as MontoBruto from V_DetalleN where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " group by Rubro) M on A.Rubro = M.Rubro " & _
                    "order by A.Orden2, A.Rubro"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 2
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraDetalleGrupoATVN(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            sql = "select A.Rubro, BrutoAcum - isnull(MontoBruto, 0) as BrutoAnt, -1 * (IGVAcum - isnull(IGV, 0)) as IGVAnt, NetoAcum - isnull(MontoNeto, 0) as NetoAnt, -1 * isnull(IGV, 0) as IGV, isnull(MontoNeto, 0) as MontoNeto " & _
                    "from (select Orden2, Rubro, sum(MontoBruto) as BrutoAcum, sum(IGV) as IGVAcum, sum(MontoNeto) as NetoAcum " & _
                    "from V_DetalleN where CodAuxiliar in ('GPATV1', 'GPATV2', 'GPATV3') and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden2, Rubro) A " & _
                    "left join (select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV, sum(Signo * MontoBruto) as MontoBruto from V_DetalleN where CodAuxiliar in ('GPATV1', 'GPATV2', 'GPATV3') and IdPeriodo = " & IdPeriodo & " group by Rubro) M on A.Rubro = M.Rubro " & _
                    "order by A.Orden2, A.Rubro"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 5
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

End Module

Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Configuration
Imports System.Data.SqlClient
'Imports System.Web

Module FuncionesFC

    Function LlenaFCExp() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.Orden as OrdenC, case C.Signo when 1 then '+' when -1 then '-' else '' end as SignoC, " & _
                "CtaFlujoRep, CodSeccion, D.Orden as OrdenD, " & _
                "CtaDetalleRep, case D.Signo when 1 then '+' when -1 then '-' else '' end as SignoD, " & _
                "Cuenta2 + '[' + convert(varchar, CodCuenta2) + ']' as Cuenta2 " & _
                "from CtaFlujoRep C left join CtaDetalleRep D on C.IdCtaFlujoRep = D.IdCtaFlujoRep left join V_CuentaFC CC on D.CodCtaOrigen = CC.CodCuenta2 " & _
                "order by C.Orden, D.Orden, Cuenta2 "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaMontoNeto(ByVal IdPeriodo As String, ByVal CodCuentaFC2 As String) As Double
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim MontoNeto As Double = 0

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "select sum(MontoBaseUSD) as MontoNeto from V_Movimiento " & _
                "where IdPeriodo = " & IdPeriodo & " and CodCuentaFC2 = " + CodCuentaFC2
        cmd = New OleDbCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() And Not IsDBNull(rdr("MontoNeto")) Then MontoNeto = Convert.ToDouble(rdr("MontoNeto"))
        rdr.Close()

        cn.Close()

        Return MontoNeto
    End Function

    Function BuscaSaldoFinalContab(ByVal IdPeriodo As String) As Double
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim SaldoFinal As Double

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "select sum(Saldo) as Saldo " & _
                "from V_SaldoBancos where IdAño * 100 + IdMes = " & IdPeriodo & " and NroCuenta like '10%' "
        cmd = New OleDbCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() And Not IsDBNull(rdr("Saldo")) Then SaldoFinal = Convert.ToDouble(rdr("Saldo"))
        rdr.Close()

        cn.Close()

        Return SaldoFinal
    End Function

    Function BuscaITFX(ByVal IdPeriodoIni As String, ByVal IdPeriodoFin As String) As Double
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim ITF As Double = 0

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "select sum(isnull(DebeUSD, 0) - isnull(HaberUSD, 0)) as MontoUSD " & _
                "from LibroMayor where NroCuenta = '954000104' and IdAño * 100 + IdMes between " & IdPeriodoIni & " and " & IdPeriodoFin
        cmd = New OleDbCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() And Not IsDBNull(rdr("MontoUSD")) Then ITF = Convert.ToDouble(rdr("MontoUSD"))
        rdr.Close()

        cn.Close()

        Return ITF
    End Function

    Sub LlenaITF()
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "delete from Ajuste where CodCuentaFC2 = 902"
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "insert into Ajuste(IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) " & _
                "select IdAño * 100 + IdMes as IdPeriodo, convert(smalldatetime, convert(varchar(10), IdAño * 10000 + IdMes * 100 + 1)) as Fecha, " & _
                "902 as CodCuentaFC2, sum(isnull(DebeUSD, 0) - isnull(HaberUSD, 0)) as MontoUSD " & _
                "from LibroMayor where NroCuenta = '954000104' " & _
                "group by IdAño * 100 + IdMes, convert(smalldatetime, convert(varchar(10), IdAño * 10000 + IdMes * 100 + 1))"
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "update Ajuste set MontoUSD = MontoUSD - ITFTesoreria " & _
                "from Ajuste A, (select IdPeriodo, sum(MontoUSD) as ITFTesoreria from Movimiento " & _
                "where CodCuentaFC = 96 group by IdPeriodo) M " & _
                "where A.IdPeriodo = M.IdPeriodo and A.CodCuentaFC2 = 902"
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub LlenaDetalle(ByVal FlagContab As Boolean, ByVal FlagNac As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "truncate table Detalle"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "insert into Detalle(IdPeriodo, Orden, CodSeccion, CodAuxiliar, Area, Orden2, Rubro, Signo, MontoNeto, IGV, MontoBruto) " & _
                "select IdPeriodo, Orden, CodSeccion, CodAuxiliar, Area, Orden2, Rubro, Signo, MontoNeto, IGV, MontoBruto "
        If FlagNac Then
            If FlagContab Then sql = sql & "from V_Detalle" Else sql = sql & "from V_DetalleT"
        Else
            If FlagContab Then sql = sql & "from V_Detalle_Ext" Else sql = sql & "from V_DetalleT_Ext"
        End If

        cmd = New SqlCommand(sql, cn)
        cmd.CommandTimeout = 300
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub Titulos(ByVal hoja As Excel.Worksheet, ByVal IdMes As Integer, ByVal fil_tit As Integer)
        If IdMes > 1 Then hoja.Cells(fil_tit, 3).Value = "ACUMULADO " & BuscaMes((IdMes - 1).ToString)
        hoja.Cells(fil_tit, 4).Value = BuscaMes(IdMes.ToString) & " NETO"
        hoja.Cells(fil_tit, 6).Value = "ACUMULADO " & BuscaMes(IdMes.ToString)
    End Sub

    Sub Titulos2(ByVal hoja As Excel.Worksheet, ByVal IdMes As Integer, ByVal fil_tit As Integer)
        If IdMes > 1 Then hoja.Cells(fil_tit, 3).Value = "ACUMULADO " & BuscaMes((IdMes - 1).ToString)
        hoja.Cells(fil_tit, 4).Value = BuscaMes(IdMes.ToString) & " NETO"
        hoja.Cells(fil_tit, 5).Value = "ACUMULADO " & BuscaMes(IdMes.ToString)
    End Sub

    Sub Titulos3(ByVal hoja As Excel.Worksheet, ByVal IdMes As Integer, ByVal fil_tit As Integer)
        If IdMes > 1 Then hoja.Cells(fil_tit, 3).Value = "ACUMULADO " & BuscaMes((IdMes - 1).ToString)
        hoja.Cells(fil_tit, 4).Value = BuscaMes(IdMes.ToString)
        hoja.Cells(fil_tit, 5).Value = "ACUMULADO " & BuscaMes(IdMes.ToString)
    End Sub

    Sub GeneraReportesFC(ByVal FlagContab As Boolean, ByVal FlagNac As Boolean, ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        'Try
        LlenaITF()
        LlenaDetalle(FlagContab, FlagNac)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("CARATULA")
        hoja.Cells(47, 2).Value = IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01"

        hoja = reporte.Worksheets("INDICE")
        hoja.Cells(1, 4).Value = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        hoja = reporte.Worksheets("FLUJO CAJA")
        CreaReporteFlujoCaja(FlagContab, hoja, IdPeriodo)
        hoja = reporte.Worksheets("MENSUAL")
        CreaReporteFlujoCajaMensual(FlagContab, hoja, IdPeriodo)
        hoja = reporte.Worksheets("F 4-INGRESOS")
        CreaReporteIngresos(hoja, IdPeriodo)
        hoja = reporte.Worksheets("F 5-POSIC CAJA")
        CreaReportePosicionCaja(FlagContab, hoja, IdPeriodo)
        hoja = reporte.Worksheets("AGENCIAS_IGV")
        CreaReporteAgenciasIGV(hoja, IdPeriodo)
        hoja = reporte.Worksheets("PAGO IMPUESTOS")
        CreaReportePagoImpuestos(hoja, IdPeriodo)
        hoja = reporte.Worksheets("RESUMEN DEUDA CORRIENTE")
        CreaReporteResumenDeudaCorriente(hoja, IdPeriodo)
        hoja = reporte.Worksheets("DETALLE DEUDA CORRIENTE")
        CreaReporteDetalleDeudaCorriente(hoja, IdPeriodo)
        hoja = reporte.Worksheets("RESUMEN PAGOS INTEREMP")
        CreaReporteResumenPagosInterEmp(hoja, IdPeriodo)
        hoja = reporte.Worksheets("PAGOS INTEREMPRESAS")
        CreaReporteDetallePagosInterEmp(hoja, IdPeriodo)
        hoja = reporte.Worksheets("RESUMEN DEUDA INDECOPI")
        CreaReporteResumenObligIndecopi(hoja, IdPeriodo)
        hoja = reporte.Worksheets("DET OBLIG INDECOPI")
        CreaReporteDetalleObligIndecopi(hoja, IdPeriodo)
        hoja = reporte.Worksheets("INVERSIONES")
        CreaReporteInversiones(hoja, IdPeriodo)

        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(reporte)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        hoja = Nothing
        reporte = Nothing
        excel = Nothing
        GC.Collect()
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub CreaReporteFlujoCaja(ByVal FlagContab As Boolean, ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño, IdMes As Integer
        Dim IdPeriodoAnt As String
        Try
            hoja.Activate()
            hoja.Cells(4, 2).Value = ("Flujo de Caja a " + BuscaPeriodo(IdPeriodo)).ToUpper()
            IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
            IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            If IdMes > 1 Then IdPeriodoAnt = (IdAño * 100 + IdMes - 1).ToString Else IdPeriodoAnt = ((IdAño - 1) * 100 + 12).ToString

            Titulos3(hoja, IdMes, 6)

            If FlagContab Then
                hoja.Cells(7, 3).Value = BuscaSaldoFinalContab(((IdAño - 1) * 100 + 12).ToString)
            Else
                hoja.Cells(7, 3).Value = BuscaCierreCaja(Convert.ToDateTime((IdAño - 1).ToString & "-12-31"))
            End If
            'hoja.Cells(7, 4).Value = BuscaSaldoFinal(FlagContab, IdPeriodoAnt)

            GeneraResumen2(hoja, IdPeriodo, "INGRES", 10, 12, 2)
            GeneraResumen2(hoja, IdPeriodo, "DESING", 14, 15, 2)
            GeneraResumen2(hoja, IdPeriodo, "EGRCOR", 19, 32, 2)
            GeneraResumen2(hoja, IdPeriodo, "GRPATV", 35, 66, 2)
            GeneraResumen2(hoja, IdPeriodo, "OBLIND", 69, 74, 2)
            GeneraResumen2(hoja, IdPeriodo, "INVERS", 77, 77, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteFlujoCajaMensual(ByVal FlagContab As Boolean, ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño, IdMes As Integer
        Dim IdPeriodoAux As String
        Try
            hoja.Activate()
            IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
            IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            For i = 1 To IdMes
                IdPeriodoAux = (IdAño * 100 + i).ToString

                'If i > 1 Then IdPeriodoAnt = (IdAño * 100 + i - 1).ToString Else IdPeriodoAnt = ((IdAño - 1) * 100 + 12).ToString
                If i = 1 Then
                    If FlagContab Then hoja.Cells(9, 2 + i).Value = BuscaSaldoFinalContab(((IdAño - 1) * 100 + 12).ToString) Else hoja.Cells(9, 2 + i).Value = BuscaCierreCaja(Convert.ToDateTime((IdAño - 1).ToString & "-12-31"))
                End If
                GeneraResumen3(hoja, IdPeriodoAux, "INGRES", 12, 14, 2)
                GeneraResumen3(hoja, IdPeriodoAux, "DESING", 16, 17, 2)
                GeneraResumen3(hoja, IdPeriodoAux, "EGRCOR", 21, 34, 2)
                GeneraResumen3(hoja, IdPeriodoAux, "GRPATV", 45, 76, 2)
                GeneraResumen3(hoja, IdPeriodoAux, "OBLIND", 79, 84, 2)
                GeneraResumen3(hoja, IdPeriodoAux, "INVERS", 87, 87, 2)
            Next i

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteIngresos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim fil_ini, fil_fin As Integer

        Try
            hoja.Activate()
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos(hoja, IdMes, 6)

            GeneraDetalle(hoja, IdPeriodo, "INGPUB", 7, 20, 2)

            fil_ini = 27 : fil_fin = 406
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "1", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "2", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "3", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "4", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "5", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "6", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "7", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "8", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "9", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "10", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "11", 1, fil_ini, 2)
            GeneraDetalleProved(hoja, IdPeriodo, "INGOTR", "15", 20, fil_ini, 2)

            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            GeneraDetalle(hoja, IdPeriodo, "INGLLA", 413, 414, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReportePosicionCaja(ByVal FlagContab As Boolean, ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim UltDiaMesAct, UltDiaMesAnt As Date
        Dim fil_ini, fil_fin As Integer

        Try
            hoja.Activate()

            UltDiaMesAnt = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddDays(-1)
            UltDiaMesAct = Convert.ToDateTime(UltDiaMesAnt).AddMonths(1)

            hoja.Cells(8, 2).Value = "POSICIÓN CAJA - BANCOS - AL " & UltDiaMesAct.ToString("dd.MM.yy")
            hoja.Cells(10, 4).Value = "Al " & UltDiaMesAnt.ToString("dd.MM.yy")
            hoja.Cells(10, 5).Value = "Al " & UltDiaMesAct.ToString("dd.MM.yy")

            fil_ini = 12 : fil_fin = 36
            GeneraPosicionCaja(FlagContab, hoja, IdPeriodo, fil_ini, fil_fin, 2)

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteAgenciasIGV(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos(hoja, IdMes, 6)

            GeneraDetalle(hoja, IdPeriodo, "AGNCIA", 7, 60, 2)
            GeneraDetalle2(hoja, IdPeriodo, "IGVAGN", 67, 67, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReportePagoImpuestos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim fil_ini As String

        Try
            hoja.Activate()
            Dim IdAño As Integer = Convert.ToInt32(IdPeriodo.Substring(0, 4))
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))

            fil_ini = 6
            For i = 1 To IdMes
                hoja.Cells(fil_ini + i, 3).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "201") + BuscaMontoNeto((IdAño * 100 + i).ToString, "904") 'IGV Cuenta Propia
                hoja.Cells(fil_ini + i, 5).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "204") + BuscaMontoNeto((IdAño * 100 + i).ToString, "906") 'Renta 3ra
                hoja.Cells(fil_ini + i, 6).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "205") + BuscaMontoNeto((IdAño * 100 + i).ToString, "907") 'Renta 4ra
                hoja.Cells(fil_ini + i, 7).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "206") + BuscaMontoNeto((IdAño * 100 + i).ToString, "908") 'Renta 5ta
                hoja.Cells(fil_ini + i, 8).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "202") + BuscaMontoNeto((IdAño * 100 + i).ToString, "905") 'IGV No Domiciliados
                hoja.Cells(fil_ini + i, 9).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "207") + BuscaMontoNeto((IdAño * 100 + i).ToString, "909") 'Renta No Domiciliados
                hoja.Cells(fil_ini + i, 10).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "208") + BuscaMontoNeto((IdAño * 100 + i).ToString, "910") 'Retención Proveedores
                hoja.Cells(fil_ini + i, 11).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "96") 'ITF
                hoja.Cells(fil_ini + i, 13).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "3") + BuscaMontoNeto((IdAño * 100 + i).ToString, "322") 'ESSALUD
                hoja.Cells(fil_ini + i, 14).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "4") + BuscaMontoNeto((IdAño * 100 + i).ToString, "323") 'ONP
                hoja.Cells(fil_ini + i, 15).Value = BuscaMontoNeto((IdAño * 100 + i).ToString, "5") + BuscaMontoNeto((IdAño * 100 + i).ToString, "935") 'ESSALUD VIDA
            Next

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteResumenDeudaCorriente(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos(hoja, IdMes, 10)
            GeneraResumen(hoja, IdPeriodo, "EGRCOR", 11, 25, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteDetalleDeudaCorriente(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim fil_ini, fil_fin As Integer
        Try
            hoja.Activate()
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos2(hoja, IdMes, 8)

            GeneraDetalle2(hoja, IdPeriodo, "BANCOS", 9, 13, 2)
            GeneraDetalle2(hoja, IdPeriodo, "SUNAT", 21, 39, 2)
            GeneraDetalle2(hoja, IdPeriodo, "UTILID", 47, 51, 2)

            fil_ini = 81 : fil_fin = 580
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "7", 60, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 581
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "8", 1, fil_ini, 2)

            fil_ini = 590 : fil_fin = 690
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "2", 60, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 699 : fil_fin = 899
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "4", 60, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 908 : fil_fin = 988
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "3", 40, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 995 : fil_fin = 1025
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "1", 20, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 1034 : fil_fin = 1084
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "5", 30, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 1091 : fil_fin = 1187
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved(hoja, IdPeriodo, "PROVED", "6", 30, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 1198 : fil_fin = 1347
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProgNac(hoja, IdPeriodo, "PRGNAC", fil_ini, fil_fin, 2)

            GeneraDetalle(hoja, IdPeriodo, "FILIAL", 1354, 1357, 2)
            GeneraDetalle2(hoja, IdPeriodo, "AFP", 1364, 1369, 2)
            GeneraDetalle2(hoja, IdPeriodo, "MTC", 1376, 1378, 2)
            GeneraDetalle(hoja, IdPeriodo, "REMUNE", 1388, 1402, 2)

            fil_ini = 1409 : fil_fin = 1418
            hoja.Rows(fil_ini.ToString & ":" & fil_fin).EntireRow.Hidden = False
            GeneraDetalleProved2(hoja, IdPeriodo, "MUNIC", "1", 3, fil_ini, 2)
            If fil_fin >= fil_ini Then hoja.Rows((fil_ini).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            GeneraDetalle(hoja, IdPeriodo, "MATFIL", 1425, 1454, 2)
            GeneraDetalle2(hoja, IdPeriodo, "OTI", 1461, 1463, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteResumenPagosInterEmp(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos(hoja, IdMes, 9)

            GeneraResumen(hoja, IdPeriodo, "GRPATV", 10, 44, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteDetallePagosInterEmp(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr, rdr2 As OleDbDataReader
        Dim sql As String
        Dim fil_ini, fil_fin, col_ini As Integer
        Dim i, j, k, i2 As Integer

        Dim IdAño, IdMes As String
        Dim CodSeccion As String

        Try
            IdAño = IdPeriodo.Substring(0, 4)
            IdMes = IdPeriodo.Substring(4, 2)
            CodSeccion = "GRPATV"
            Titulos(hoja, Convert.ToInt32(IdMes), 6)

            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            fil_ini = 7
            col_ini = 2
            fil_fin = 21

            sql = "select ltrim(F1.Rubro) as Rubro, MontoBrutoA, MontoNeto, IGV " & _
                    "from (select Orden, Rubro, sum(MontoBruto) as MontoBrutoA from V_Flujo " & _
                    "where CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo < " & IdPeriodo & " " & _
                    "and ltrim(Rubro) in (select Empresa from GrupoATV where Nivel = 2) group by Orden, Rubro) F1, " & _
                    "(select Rubro, sum(MontoNeto) as MontoNeto, sum(IGV) as IGV from V_Flujo " & _
                    "where CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo = " & IdPeriodo & " " & _
                    "and ltrim(Rubro) in (select Empresa from GrupoATV where Nivel = 2) group by Rubro) F2 " & _
                    "where F1.Rubro = F2.Rubro order by F1.Orden"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 3
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()
            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            fil_ini = 27
            fil_fin = 227
            col_ini = 2

            'sql = "select F1.Rubro, MontoBrutoA, MontoNeto, IGV " & _
            '        "from (select Orden, Rubro, sum(MontoBruto) as MontoBrutoA from V_Flujo " & _
            '        "where CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo < " & IdPeriodo & " " & _
            '        "and Rubro in (select Empresa from GrupoATV where Nivel = 1) group by Orden, Rubro) F1, " & _
            '        "(select Rubro, sum(MontoNeto) as MontoNeto, sum(IGV) as IGV from V_Flujo " & _
            '        "where CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo = " & IdPeriodo & " " & _
            '        "and Rubro in (select Empresa from GrupoATV where Nivel = 1) group by Rubro) F2 " & _
            '        "where F1.Rubro = F2.Rubro order by F1.Orden"

            'sql = "select Empresa from GrupoATV where Nivel = 1 order by Empresa"
            sql = "select G.Orden, Empresa, sum(MontoBruto) as MontoBrutoA from V_Flujo F, GrupoATV G " & _
                    "where F.Rubro = G.Empresa and CodSeccion = '" & CodSeccion & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Nivel = 1 " & _
                    "group by G.Orden, Empresa having sum(MontoBruto) > 0 order by G.Orden, Empresa"

            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            k = 0
            While rdr.Read
                k = k + 1
                hoja.Cells(fil_ini + i - 3, col_ini).Value = rdr(1).ToString.ToUpper
                sql = "select F1.Rubro, BrutoAcumAct - isnull(MontoNeto, 0) - isnull(IGV, 0) as BrutoAcumAnt, MontoNeto, IGV " & _
                        "from (select Orden2, Rubro, sum(Signo * MontoBruto) as BrutoAcumAct from Detalle " & _
                        "where CodSeccion = '" & CodSeccion & "' and Area = '" & rdr(1).ToString & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 < 99 " & _
                        "group by Orden2, Rubro) F1 left join " & _
                        "(select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV from Detalle " & _
                        "where CodSeccion = '" & CodSeccion & "' and Area = '" & rdr(1).ToString & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo = " & IdPeriodo & " " & _
                        "group by Rubro) F2 " & _
                        "on F1.Rubro = F2.Rubro order by F1.Orden2"
                cmd = New OleDbCommand(sql, cn)
                rdr2 = cmd.ExecuteReader
                i2 = i
                While rdr2.Read
                    For j = 0 To 3
                        hoja.Cells(fil_ini + i2, col_ini + j).Value = rdr2(j)
                    Next
                    i2 = i2 + 1
                End While
                If k > 1 And i2 - i < 7 Then hoja.Rows((fil_ini + i2).ToString & ":" & (fil_ini + i + 6).ToString).EntireRow.Hidden = True

                If k = 1 Then i = i + 15 Else i = i + 12
                If k = 3 Or k = 8 Or k = 13 Then i = i + 3
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i - 3 Then hoja.Rows((fil_ini + i - 3).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub CreaReporteResumenObligIndecopi(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos2(hoja, IdMes, 9)

            GeneraResumen2(hoja, IdPeriodo, "OBLIND", 10, 17, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteDetalleObligIndecopi(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos2(hoja, IdMes, 8)

            GeneraDetalle2(hoja, IdPeriodo, "ICPSUN", 9, 14, 2)
            GeneraDetalle2(hoja, IdPeriodo, "ICPPRV", 21, 24, 2)
            GeneraDetalle2(hoja, IdPeriodo, "ICPREM", 31, 32, 2)
            GeneraDetalle2(hoja, IdPeriodo, "ICPCER", 39, 40, 2)
            GeneraDetalle2(hoja, IdPeriodo, "ICPBAN", 47, 50, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub CreaReporteInversiones(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Try
            hoja.Activate()
            Dim IdMes As Integer = Convert.ToInt32(IdPeriodo.Substring(4, 2))
            Titulos(hoja, IdMes, 9)

            GeneraDetalle(hoja, IdPeriodo, "INVERS", 10, 19, 2)

            If IdPeriodo.Substring(4, 2) = "01" Then hoja.Columns("C:C").EntireColumn.Hidden = True

            'hoja.PageSetup.PrintArea = "$B$1:$L$" + CStr(fil_ini + i)
            'hoja.Cells(1, 1).Select()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub GeneraPosicionCaja(ByVal FlagContab As Boolean, ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer
        Dim IdAño, IdMes As Integer
        Dim IdPeriodoAnt As String

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
        IdPeriodoAnt = (Convert.ToInt32(IdPeriodo) - 1).ToString
        If IdMes > 1 Then IdPeriodoAnt = (IdAño * 100 + IdMes - 1).ToString Else IdPeriodoAnt = ((IdAño - 1) * 100 + 12).ToString

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select Banco, Moneda, M0.Saldo as SaldoAnt, M1.Saldo as SaldoAct, Orden from CuentaContable C left join" & _
                    "(select NroCuenta, Saldo " & _
                    "from V_SaldoBancos where IdAño * 100 + IdMes = " & IdPeriodoAnt & ") M0 on C.NroCuenta = M0.NroCuenta join " & _
                    "(select NroCuenta, Saldo " & _
                    "from V_SaldoBancos where IdAño * 100 + IdMes = " & IdPeriodo & ") M1 on M0.NroCuenta = M1.NroCuenta "
            If FlagContab Then sql = sql & "where C.NroCuenta like '10%' " Else sql = sql & "where C.NroCuenta like '104%' "
            sql = sql & "order by Orden"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 3
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

    Sub GeneraResumen(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer
        Dim IdAño As String

        'Revisado 2011-06-16
        IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select F1.Rubro, BrutoAcum - MontoBruto as BrutoAnt, MontoNeto, IGV " & _
                    "from " & _
                    "(select Orden, Orden2, Rubro, sum(Signo * isnull(MontoBruto, 0)) as BrutoAcum from V_Flujo where CodSeccion = '" & CodSeccion & "' " & _
                    "and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden, Orden2, Rubro) F1, " & _
                    "(select Rubro, sum(Signo * isnull(MontoNeto, 0)) as MontoNeto, sum(Signo * isnull(IGV, 0)) as IGV, sum(Signo * isnull(MontoBruto, 0)) as MontoBruto from V_Flujo where CodSeccion = '" & CodSeccion & "' " & _
                    "and IdPeriodo = " & IdPeriodo & " group by Rubro) F2 " & _
                    "where F1.Rubro = F2.Rubro order by F1.Orden, F1.Orden2, F1.Rubro"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 3
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

    Sub GeneraResumen2(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer
        Dim IdAño As String

        'Revisado 2011-06-16
        IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select F1.Rubro, BrutoAcum - MontoBruto as BrutoAnt, MontoBruto " & _
                    "from " & _
                    "(select Orden, Orden2, Rubro, sum(Signo * isnull(MontoBruto, 0)) as BrutoAcum from V_Flujo where CodSeccion = '" & CodSeccion & "' " & _
                    "and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden, Orden2, Rubro) F1, " & _
                    "(select Rubro, sum(Signo * isnull(MontoBruto, 0)) as MontoBruto from V_Flujo where CodSeccion = '" & CodSeccion & "' " & _
                    "and IdPeriodo = " & IdPeriodo & " group by Rubro) F2 " & _
                    "where F1.Rubro = F2.Rubro order by F1.Orden, F1.Orden2, F1.Rubro"
            cmd = New SqlCommand(sql, cn)
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

    Sub GeneraResumen3(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i As Integer
        Dim IdMes As Integer

        'Revisado 2011-06-16
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select Rubro, Signo * isnull(MontoBruto, 0) as MontoBruto from V_Flujo where CodSeccion = '" & CodSeccion & "' " & _
                    "and IdPeriodo = " & IdPeriodo & " order by Orden, Orden2, Rubro"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                hoja.Cells(fil_ini + i, col_ini).Value = rdr("Rubro")
                hoja.Cells(fil_ini + i, col_ini + IdMes).Value = rdr("MontoBruto")
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

    Sub GeneraDetalle(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select A.Rubro, MontoAcum - isnull(MontoNeto, 0) - isnull(IGV, 0) as MontoBrutoA, isnull(MontoNeto, 0) as MontoNeto, isnull(IGV, 0) as IGV " & _
                    "from (select Orden2, Rubro, sum(Signo * MontoBruto) as MontoAcum " & _
                    "from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden2, Rubro) A " & _
                    "left join (select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " group by Rubro) M on A.Rubro = M.Rubro " & _
                    "order by A.Orden2, A.Rubro"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                If rdr(0).ToString.Substring(0, 1) <> " " Then
                    If (rdr("rubro").ToString <> "Sub Total US$") Then
                        For j = 0 To 3
                            hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                        Next
                    End If
                    i = i + 1
                End If
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraDetalle2(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select A.Rubro, MontoAcum - isnull(MontoBruto, 0) as MontoBrutoA, isnull(MontoBruto, 0) as MontoBruto " & _
                    "from (select Orden2, Rubro, sum(MontoBruto) as MontoAcum " & _
                    "from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Orden2, Rubro) A " & _
                    "left join (select Rubro, MontoBruto from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & ") M on A.Rubro = M.Rubro " & _
                    "order by A.Orden2, A.Rubro"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                If rdr(0).ToString.Substring(0, 1) <> " " Then
                    For j = 0 To 2
                        hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                    Next
                    i = i + 1
                End If
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraDetalleProved(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal Orden2 As String, ByVal Top As Integer, ByRef fil_ini As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j, filV_ini, filV_fin As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            i = 0

            sql = "select A.Rubro, MontoAcum - isnull(MontoNeto, 0) - isnull(IGV, 0) as MontoBrutoA, isnull(MontoNeto, 0) as MontoNeto, isnull(IGV, 0) as IGV " & _
                    "from (select top " & Top.ToString & " Rubro, sum(Signo * MontoBruto) as MontoAcum from Detalle " & _
                    "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro order by 2 desc) A " & _
                    "left join (select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro) M on A.Rubro = M.Rubro " & _
                    "order by A.Rubro "

            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader
            While rdr.Read
                For j = 0 To 3
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()

            If (Top > 1) And i = Top Then
                filV_ini = fil_ini + i + 1
                hoja.Cells(filV_ini - 1, col_ini).Value = "VARIOS PROVEEDORES"

                sql = "select A.Rubro, MontoAcum - isnull(MontoNeto, 0) - isnull(IGV, 0) as MontoBrutoA, isnull(MontoNeto, 0) as MontoNeto, isnull(IGV, 0) as IGV " & _
                        "from (select Rubro, sum(Signo * MontoBruto) as MontoAcum from Detalle " & _
                        "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " " & _
                        "and Rubro not in (select Rubro from (select top " & Top.ToString & " Rubro, sum(Signo * MontoBruto) as MontoAcum from Detalle " & _
                        "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro order by 2 desc) T) " & _
                        "group by Rubro) A " & _
                        "left join (select Rubro, sum(Signo * MontoNeto) as MontoNeto, sum(Signo * IGV) as IGV from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro) M on A.Rubro = M.Rubro " & _
                        "order by A.Rubro "

                cmd = New SqlCommand(sql, cn)
                rdr = cmd.ExecuteReader

                i = i + 1
                While rdr.Read
                    For j = 0 To 3
                        hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                    Next
                    i = i + 1
                End While
                rdr.Close()
                filV_fin = fil_ini + i - 1
                hoja.Cells(filV_ini - 1, col_ini + 1).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Cells(filV_ini - 1, col_ini + 2).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Cells(filV_ini - 1, col_ini + 3).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Rows((filV_ini).ToString & ":" & filV_fin.ToString).EntireRow.Hidden = True
            Else
                filV_ini = fil_ini + i
            End If

            'If Top > 1 Then
            '    hoja.Cells(fil_ini, col_ini + 1).FormulaR1C1 = "=SUM(R" & (fil_ini + 1).ToString & "C[0]:R" & (filV_ini - 1).ToString & "C[0])"
            '    hoja.Cells(fil_ini, col_ini + 2).FormulaR1C1 = "=SUM(R" & (fil_ini + 1).ToString & "C[0]:R" & (filV_ini - 1).ToString & "C[0])"
            '    hoja.Cells(fil_ini, col_ini + 3).FormulaR1C1 = "=SUM(R" & (fil_ini + 1).ToString & "C[0]:R" & (filV_ini - 1).ToString & "C[0])"
            'End If

            If Top > 1 And i = Top Then
                fil_ini = filV_fin + 1
            Else
                fil_ini = fil_ini + i
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraDetalleProved2(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal Orden2 As String, ByVal Top As String, ByRef fil_ini As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j, filV_ini, filV_fin As Integer

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "select A.Rubro, MontoAcum - isnull(MontoBruto, 0) as MontoBrutoA, isnull(MontoBruto, 0) as MontoBruto " & _
                    "from (select top " & Top & " Rubro, sum(MontoBruto) as MontoAcum from Detalle " & _
                    "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro order by 2 desc) A " & _
                    "left join (select Rubro, MontoBruto from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " and Orden2 = " & Orden2 & ") M on A.Rubro = M.Rubro " & _
                    "order by A.Rubro "

            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                For j = 0 To 2
                    hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                Next
                i = i + 1
            End While
            rdr.Close()

            If i = Convert.ToInt32(Top) Then
                filV_ini = fil_ini + i + 1
                hoja.Cells(filV_ini - 1, col_ini).Value = "VARIOS PROVEEDORES"

                sql = "select A.Rubro, MontoAcum - isnull(MontoBruto, 0) as MontoBrutoA, isnull(MontoBruto, 0) as MontoBruto " & _
                        "from (select Rubro, sum(MontoBruto) as MontoAcum from Detalle " & _
                        "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " " & _
                        "and Rubro not in (select Rubro from (select top " & Top & " Rubro, sum(MontoBruto) as MontoAcum from Detalle " & _
                        "where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " and Orden2 = " & Orden2 & " group by Rubro order by 2 desc) T) " & _
                        "group by Rubro) A " & _
                        "left join (select Rubro, MontoBruto from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & " and Orden2 = " & Orden2 & ") M on A.Rubro = M.Rubro " & _
                        "order by A.Rubro "

                cmd = New SqlCommand(sql, cn)
                rdr = cmd.ExecuteReader

                i = i + 1
                While rdr.Read
                    For j = 0 To 2
                        hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                    Next
                    i = i + 1
                End While
                rdr.Close()
                filV_fin = fil_ini + i - 1
                hoja.Cells(filV_ini - 1, col_ini + 1).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Cells(filV_ini - 1, col_ini + 2).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Cells(filV_ini - 1, col_ini + 3).FormulaR1C1 = "=SUM(R" & filV_ini.ToString & "C[0]:R" & filV_fin & "C[0])"
                hoja.Rows((filV_ini).ToString & ":" & filV_fin.ToString).EntireRow.Hidden = True
                fil_ini = filV_fin + 1
            Else
                fil_ini = fil_ini + i
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraDetalleProgNac(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodAuxiliar As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim i, j As Integer
        Dim aux As String

        Dim IdAño = IdPeriodo.Substring(0, 4)

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            aux = "="
            sql = "select A.Rubro, MontoAcum - isnull(MontoNeto, 0) - isnull(IGV, 0) as MontoBrutoA, isnull(MontoNeto, 0) as MontoNeto, isnull(IGV, 0) as IGV, A.Area, A.Orden2 " & _
                    "from (select Area, Orden2, Rubro, sum(MontoBruto) as MontoAcum " & _
                    "from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo / 100 = " & IdAño & " and IdPeriodo <= " & IdPeriodo & " group by Area, Orden2, Rubro) A " & _
                    "left join (select Area, Orden2, Rubro, MontoNeto, IGV from Detalle where CodAuxiliar = '" & CodAuxiliar & "' and IdPeriodo = " & IdPeriodo & ") M on A.Area = M.Area and A.Orden2 = M.Orden2 " & _
                    "order by A.Area, A.Orden2"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                If rdr(0).ToString.Substring(0, 1) <> " " Then
                    For j = 0 To 3
                        hoja.Cells(fil_ini + i, col_ini + j).Value = rdr(j)
                    Next
                    If Convert.ToInt32(rdr("Orden2")) = 0 Then
                        aux = aux + " R" & (fil_ini + i) & "C[0] + "
                    Else
                        hoja.Rows((fil_ini + i).ToString & ":" & (fil_ini + i).ToString).EntireRow.Hidden = True
                    End If
                    i = i + 1
                End If
            End While
            rdr.Close()

            If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

            For j = 1 To 4
                hoja.Cells(fil_fin + 1, col_ini + j).FormulaR1C1 = aux.Substring(0, aux.Length - 2)
            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

End Module

Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data.SqlClient

Module FuncionesFCS

    Sub GeneraReportesFCS(ByVal FlagNac As Boolean, ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal FechaIni As Date, ByVal FechaFin As Date)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet
        Dim IdAño, IdPeriodoAct, IdPeriodoAnt, IdPeriodoAnt2, IdPeriodoEne, IdPeriodoAct2, IdPeriodoAct3 As Integer

        Try
            excel = New Excel.Application
            excel.DisplayAlerts = False
            reporte = excel.Workbooks.Open(RutaPlantilla)

            hoja = reporte.Worksheets("Flujo ATV")

            hoja.Cells(6, 3).Value = "A " & FechaFin.AddMonths(-2).ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(6, 4).Value = FechaFin.AddMonths(-1).ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            If FechaIni.ToString("dd") = "01" Then
                hoja.Cells(6, 5).Value = "'01 a " & FechaFin.ToString("dd") & " " & FechaFin.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
                hoja.Cells(6, 6).Value = ""
            Else
                hoja.Cells(6, 5).Value = "'01 a " & FechaIni.AddDays(-1).ToString("dd") & " " & FechaIni.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
                hoja.Cells(6, 6).Value = FechaIni.ToString("dd") & " a " & FechaFin.ToString("dd") & " " & FechaFin.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            End If
            hoja.Cells(6, 7).Value = "ENE - " & FechaFin.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(6, 8).Value = FechaFin.ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(6, 9).Value = FechaFin.AddMonths(1).ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(6, 10).Value = FechaFin.AddMonths(2).ToString("MMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

            IdAño = Convert.ToInt32(FechaFin.ToString("yyyy"))
            IdPeriodoAct = Convert.ToInt32(FechaFin.ToString("yyyyMM"))
            IdPeriodoAnt = Convert.ToInt32(FechaFin.AddMonths(-1).ToString("yyyyMM"))
            IdPeriodoAnt2 = Convert.ToInt32(FechaFin.AddMonths(-2).ToString("yyyyMM"))
            IdPeriodoEne = Convert.ToInt32(FechaFin.AddMonths(-2).ToString("yyyy") & "01")
            IdPeriodoAct2 = Convert.ToInt32(FechaFin.AddMonths(1).ToString("yyyyMM"))
            IdPeriodoAct3 = Convert.ToInt32(FechaFin.AddMonths(2).ToString("yyyyMM"))

            hoja.Cells(8, 3).Value = BuscaCierreCaja(Convert.ToDateTime((IdAño - 1).ToString & "-12-31")) 'FuncionesRep.BuscaSaldoFinal(False, ((IdAño - 1) * 100 + 12).ToString)

            CreaDetalleS(FlagNac, IdPeriodoEne, IdPeriodoAnt2, IdAño.ToString & "-01-01", FechaFin)
            GeneraMes(hoja, "E", IdPeriodoAnt2, 3)
            CreaDetalleS(FlagNac, IdPeriodoAnt, IdPeriodoAnt, IdAño.ToString & "-01-01", FechaFin)
            GeneraMes(hoja, "E", IdPeriodoAnt, 4)

            If FechaIni.ToString("dd") = "01" Then
                CreaDetalleS(FlagNac, IdPeriodoAct, IdPeriodoAct, FechaIni, FechaFin)
                GeneraMes(hoja, "E", IdPeriodoAct, 5)
                hoja.Columns("F:F").EntireColumn.Hidden = True
            Else
                CreaDetalleS(FlagNac, IdPeriodoAct, IdPeriodoAct, IdAño.ToString & "-01-01", FechaIni.AddDays(-1).ToString("yyyy-MM-dd"))
                GeneraMes(hoja, "E", IdPeriodoAct, 5)
                CreaDetalleS(FlagNac, IdPeriodoAct, IdPeriodoAct, FechaIni, FechaFin)
                GeneraMes(hoja, "E", IdPeriodoAct, 6)
            End If
            GeneraMes(hoja, "P", IdPeriodoAct, 8)
            GeneraMes(hoja, "P", IdPeriodoAct2, 9)
            GeneraMes(hoja, "P", IdPeriodoAct3, 10)

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
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
        End Try
    End Sub

    Sub CreaDetalleS(ByVal FlagNac As Boolean, ByVal IdPeriodoIni As Integer, ByVal IdPeriodoFin As Integer, ByVal FechaIni As String, ByVal FechaFin As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            If FlagNac Then sql = "sp_Crea_DetalleS" Else sql = "sp_Crea_DetalleS_Ext"
            cmd = New OleDbCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("IdPeriodoIni", IdPeriodoIni)
            cmd.Parameters.AddWithValue("IdPeriodoFin", IdPeriodoFin)
            cmd.Parameters.AddWithValue("FechaIni", FechaIni)
            cmd.Parameters.AddWithValue("FechaFin", FechaFin)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraMes(ByVal hoja As Excel.Worksheet, ByVal CodVersion As String, ByVal IdPeriodo As Integer, ByVal col As Integer)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "INGRES", 9, 13, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "PROVED", 16, 35, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "REMUNE", 38, 49, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "IMPTOS", 52, 63, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "INVERS", 66, 70, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "MATFIL", 73, 80, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "GRPATV", 83, 102, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "PATADO", 105, 105, col)
        GeneraDetalle(hoja, CodVersion, IdPeriodo, "INDECO", 107, 111, col)
    End Sub

    Sub GeneraDetalle(ByVal hoja As Excel.Worksheet, ByVal CodVersion As String, ByVal IdPeriodo As Integer, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i As Integer

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            sql = "select F.Rubro, D.MontoBruto from DetalleS D right join V_FlujoS F on D.CodVersion = F.CodVersion and D.IdPeriodo = F.IdPeriodo and D.CodDetalle = F.CodDetalle and D.Rubro = F.Rubro "
            sql &= "where F.CodVersion = '" & CodVersion & "' and F.IdPeriodo = " & IdPeriodo.ToString & " and F.CodSeccion = '" & CodSeccion & "' order by Rubro"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")
                hoja.Cells(fil_ini + i, col).Value = rdr("MontoBruto")
                'If CodSeccion <> "IMPTOS" Or rdr("Rubro") <> "ITF" Then
                '    hoja.Cells(fil_ini + i, col).Value = rdr("MontoBruto")
                'Else
                '    Dim IdPeriodoEne As String = IdPeriodo.ToString.Substring(0, 4) & "01"
                '    Dim ITF As Double
                '    If col = 3 Then ITF = FuncionesRep.BuscaITF(IdPeriodoEne, IdPeriodo.ToString) Else ITF = FuncionesRep.BuscaITF(IdPeriodo.ToString, IdPeriodo.ToString)
                '    If ITF > 0 Then
                '        hoja.Cells(fil_ini + i, col).Value = ITF
                '    Else
                '        hoja.Cells(fil_ini + i, col).Value = rdr("MontoBruto")
                '    End If
                'End If
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

    Function BuscaTipoCambioCierre(ByVal FechaCierre As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim TipoCambio As Double

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select top 1 isnull(TipoCambio, 0) as TipoCambio from CierreCaja where FechaCierre = '" & FechaCierre & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then TipoCambio = Convert.ToDouble(rdr("TipoCambio")) Else TipoCambio = 0
        rdr.Close()

        cn.Close()

        Return TipoCambio
    End Function

    Function BuscaCierreCaja(ByVal FechaCierre As Date) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim MontoSoles, MontoDolares, TipoCambio, MontoCierre As Double

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(sum(MontoCierre), 0) as MontoSoles from CierreCaja where CodMoneda = 1 and FechaCierre = '" & FechaCierre.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        MontoSoles = Convert.ToDouble(rdr("MontoSoles"))
        rdr.Close()

        sql = "select isnull(sum(MontoCierre), 0) as MontoDolares from CierreCaja where CodMoneda = 2 and FechaCierre = '" & FechaCierre.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        MontoDolares = rdr("MontoDolares").ToString()
        rdr.Close()

        TipoCambio = 0
        sql = "select top 1 TipoCambio from CierreCaja where FechaCierre = '" & FechaCierre & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then TipoCambio = Convert.ToDouble(rdr("TipoCambio"))
        rdr.Close()

        If TipoCambio > 0 Then MontoCierre = MontoSoles / TipoCambio + MontoDolares Else MontoCierre = 0

        cn.Close()
        Return MontoCierre
    End Function

    Sub BuscaDatosCierreCaja(ByVal FechaCierre As Date, ByRef MontoSoles As Double, ByVal MontoDolares As Double, ByRef TipoCambio As Double)
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select sum(MontoCierre) as MontoSoles from CierreCaja where CodMoneda = 1 and FechaCierre = '" & FechaCierre.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        MontoSoles = Convert.ToDouble(rdr("MontoSoles"))
        rdr.Close()

        sql = "select sum(MontoCierre) as MontoDolares from CierreCaja where CodMoneda = 2 and FechaCierre = '" & FechaCierre.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        MontoDolares = rdr("MontoDolares").ToString()
        rdr.Close()

        sql = "select top 1 TipoCambio from CierreCaja where FechaCierre = '" & FechaCierre.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        TipoCambio = Convert.ToDouble(rdr("TipoCambio"))
        rdr.Close()

        cn.Close()
    End Sub

    Sub BuscaMovimientos(ByVal Fecha As Date, ByRef MontoSoles As Double, ByRef MontoDolares As Double, ByRef DifCambio As Double)
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(sum(case when Cuenta like 'ING-%' or Cuenta like 'OIN-%' then Monto else -1 * Monto end), 0) as MontoSoles, " & _
                "isnull(sum(case when Cuenta like 'ING-%' or Cuenta like 'OIN-%' then MontoUSD else -1 * MontoUSD end), 0) as MontoDolares " & _
                "from Movimiento M, CuentaFC C " & _
                "where M.CodCuentaFC = C.CodCuenta and CodMoneda = 1 and Fecha = '" & Fecha.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        MontoSoles = Convert.ToDouble(rdr("MontoSoles"))
        MontoDolares = Convert.ToDouble(rdr("MontoDolares"))
        rdr.Close()

        '99	OIN-COMPRA DOLAR-INGRESO
        '100 OEG-COMPRA DOLAR -EGRESO
        '101 OIN-VENTA DOLAR INGRESO
        '102 OEG-VENTA DOLAR EGRESO
        sql = "select isnull(sum(case when CodCuentaFC in (99, 101) then MontoUSD else -1 * MontoUSD end), 0) as DifCambio " & _
                "from Movimiento M " & _
                "where CodCuentaFC in (99, 100, 101, 102) and Fecha = '" & Fecha.ToString("yyyy-MM-dd") & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        DifCambio = Convert.ToDouble(rdr("DifCambio"))
        rdr.Close()

        cn.Close()
    End Sub

    Sub CalculaAjustesTipoCambio(ByVal Fecha As Date)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim sql As String
        Dim IdPeriodo As String
        Dim FechaAnt As Date
        Dim DifCambio, MontoSoles, MontoDolares, TipoCambio, TipoCambio2, Ajuste As Double

        If Fecha.DayOfWeek() = DayOfWeek.Saturday Or Fecha.DayOfWeek = DayOfWeek.Sunday Then Exit Sub

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        MontoSoles = 0 : MontoDolares = 0 : TipoCambio = 0

        If Fecha.DayOfWeek() = DayOfWeek.Monday Then FechaAnt = Fecha.AddDays(-3) Else FechaAnt = Fecha.AddDays(-1)

        IdPeriodo = Fecha.ToString("yyyyMM")

        TipoCambio2 = BuscaTipoCambioCierre(Fecha)
        BuscaDatosCierreCaja(FechaAnt, MontoSoles, MontoDolares, TipoCambio)

        sql = "delete from Ajuste where CodCuentaFC2 = 946 and Fecha = '" & Fecha.ToString("yyyy-MM-dd") & "' "
        GrabaLog(0, sql)
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()

        If TipoCambio > 0 And TipoCambio2 > 0 Then
            Ajuste = MontoSoles * (1 / TipoCambio2 - 1 / TipoCambio)
            sql = "insert into Ajuste(IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) " & _
                    "values(" & IdPeriodo & ",'" & Fecha.ToString("yyyy-MM-dd") & "', 946, " & Ajuste.ToString & ")"
            GrabaLog(0, sql)
            cmd = New OleDbCommand(sql, cn)
            cmd.ExecuteNonQuery()
        End If
        MontoSoles = 0 : MontoDolares = 0 : DifCambio = 0
        BuscaMovimientos(Fecha, MontoSoles, MontoDolares, DifCambio)

        sql = "delete from Ajuste where CodCuentaFC2 in (901, 928) and Fecha = '" & Fecha.ToString("yyyy-MM-dd") & "' "
        GrabaLog(0, sql)
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()

        If TipoCambio > 0 And TipoCambio2 > 0 Then Ajuste = MontoSoles / TipoCambio2 - MontoDolares Else Ajuste = 0
        If Ajuste <> 0 Then
            sql = "insert into Ajuste(IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) " & _
                    "values(" & IdPeriodo & ",'" & Fecha.ToString("yyyy-MM-dd") & "', 928, " & Ajuste.ToString & ")"
            GrabaLog(0, sql)
            cmd = New OleDbCommand(sql, cn)
            cmd.ExecuteNonQuery()
        End If
        If DifCambio <> 0 Then
            sql = "insert into Ajuste(IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) " & _
                    "values(" & IdPeriodo & ",'" & Fecha.ToString("yyyy-MM-dd") & "', 901, " & DifCambio.ToString & ")"
            GrabaLog(0, sql)
            cmd = New OleDbCommand(sql, cn)
            cmd.ExecuteNonQuery()
        End If

        cn.Close()
    End Sub

    Sub CalculaAjustesTipoCambio()
        Dim Fecha As Date
        Dim FechaFin As Date

        If Date.Now.DayOfWeek = DayOfWeek.Monday Then FechaFin = Date.Now.AddDays(-3) Else FechaFin = Date.Now.AddDays(-1)

        Fecha = Convert.ToDateTime("2011-09-01")
        While Fecha <= FechaFin
            CalculaAjustesTipoCambio(Fecha)
            Fecha = Fecha.AddDays(1)
        End While

    End Sub

    Sub GeneraReportesFCS2(ByVal FlagNac As Boolean, ByVal FlagDiario As Boolean, ByVal TipoMonto As String, ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal FechaIni As Date, ByVal FechaFin As Date)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet
        Dim rango As Excel.Range

        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String

        Dim IdPeriodoAnt, IdPeriodoAct As Integer
        Dim DiaUno As Date
        Dim tabla As String
        Dim fil_ini As Integer

        'Try
        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        If FlagDiario Then

            fil_ini = 4
            DiaUno = Convert.ToDateTime(FechaIni.ToString("yyyy-MM") & "-01")

            CreaDetalleD(FlagNac, DiaUno, FechaFin)

            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            If FlagNac Then tabla = "V_Movimiento" Else tabla = "V_Movimiento_Ext"
            sql = "select distinct Fecha from " & tabla & " " & _
                    "where Fecha >= '" & DiaUno & "' and Fecha <= '" & FechaFin & "' order by 1"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader
            While rdr.Read  ' DiaUno.AddDays(i) <= FechaFin
                GeneraFCDiario(excel, reporte, fil_ini, FlagNac, TipoMonto, Convert.ToDateTime(rdr("Fecha")))
                fil_ini = fil_ini + 76
            End While

            hoja = reporte.Worksheets("Flujo ATV")
            rango = hoja.Rows("79:79")
            hoja = reporte.Worksheets("Flujo ATV Diario")
            rango.Copy(hoja.Rows((fil_ini + 1).ToString))
            hoja.Cells(fil_ini + 1, 2).Value = "SALDO FINAL AL " & FechaFin.ToString("dd.MM.yy") & " US$"
            hoja.Cells(fil_ini + 1, 3).Value = BuscaCierreCaja(FechaFin)

            hoja = reporte.Worksheets("Flujo ATV Diario")
            hoja.Columns("D:E").EntireColumn.Hidden = True

            rdr.Close()
            cn.Close()

        End If

        If Not FlagDiario Then

            hoja = reporte.Worksheets("Flujo ATV")
            fil_ini = 4
            hoja.Cells(fil_ini + 1, 2).Value = "DEL " & FechaIni.ToString("dd") & " AL " & FechaFin.ToString("dd") & " DE " & FechaFin.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper & " " & FechaFin.ToString("yyyy")

            If TipoMonto = "MontoBruto" Then
                hoja.Cells(fil_ini, 2).Value = "FLUJO DE CAJA EJECUTADO"
                If FechaIni.DayOfWeek() <> DayOfWeek.Monday Then
                    hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & FechaIni.AddDays(-1).ToString("dd.MM.yy") & " US$"
                    hoja.Cells(fil_ini + 5, 3).Value = BuscaCierreCaja(FechaIni.AddDays(-1).ToString("yyyy-MM-dd"))
                Else
                    hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & FechaIni.AddDays(-3).ToString("dd.MM.yy") & " US$"
                    hoja.Cells(fil_ini + 5, 3).Value = BuscaCierreCaja(FechaIni.AddDays(-3).ToString("yyyy-MM-dd"))
                End If
                hoja.Cells(fil_ini + 74, 2).Value = "SALDO AL " & FechaFin.ToString("dd.MM.yy") & " US$"
                hoja.Rows((fil_ini + 75).ToString & ":" & (fil_ini + 75).ToString).EntireRow.Hidden = True
            Else
                hoja.Cells(fil_ini, 2).Value = "FLUJO DE CAJA EJECUTADO RAG"
                If FechaIni.DayOfWeek() <> DayOfWeek.Monday Then
                    hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & FechaIni.AddDays(-1).ToString("dd.MM.yy") & " US$"
                Else
                    hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & FechaIni.AddDays(-3).ToString("dd.MM.yy") & " US$"
                End If
                hoja.Cells(fil_ini + 74, 2).Value = "SALDO SIN IGV AL " & FechaFin.ToString("dd.MM.yy") & " US$"
                hoja.Cells(fil_ini + 75, 2).Value = "SALDO CON IGV AL " & FechaFin.ToString("dd.MM.yy") & " US$"
            End If

            IdPeriodoAnt = Convert.ToInt32(FechaIni.ToString("yyyyMM"))
            IdPeriodoAct = Convert.ToInt32(FechaFin.ToString("yyyyMM"))
            CreaDetalleS(FlagNac, IdPeriodoAnt, IdPeriodoAct, FechaIni, FechaFin)
            GeneraDetalleS(True, hoja, "E", IdPeriodoAct, "INGRES", TipoMonto, 10, 14, 3)
            GeneraDetalleS(True, hoja, "E", IdPeriodoAct, "EGRESO", TipoMonto, 17, 76, 3)

            hoja = reporte.Worksheets("Anexos")
            GeneraInversiones(hoja, TipoMonto, FechaIni, FechaFin, 8, 22, 3)

            hoja = reporte.Worksheets("Detalle")
            CreaIdMovimientoS(FlagNac, FechaIni, FechaFin)
            GeneraDetalleCompleto(hoja, FlagNac, FechaIni, FechaFin)

        End If

        If FlagDiario Then
            DiaUno = Convert.ToDateTime(FechaIni.ToString("yyyy-MM") & "-01")
            hoja = reporte.Worksheets("Detalle Diario")
            CreaIdMovimientoS(FlagNac, DiaUno, FechaFin)
            GeneraDetalleDiarioCompleto(hoja, FlagNac, DiaUno, FechaFin)
        End If

        If Not FlagDiario Then
            hoja = reporte.Worksheets("Flujo ATV Diario")
            hoja.Delete()
            hoja = reporte.Worksheets("Detalle Diario")
            hoja.Delete()
        Else
            hoja = reporte.Worksheets("Flujo ATV")
            hoja.Delete()
            hoja = reporte.Worksheets("Anexos")
            hoja.Delete()
            hoja = reporte.Worksheets("Detalle")
            hoja.Delete()
        End If

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

    Sub GeneraDetalleS(ByVal FlagSemanal As Boolean, ByVal hoja As Excel.Worksheet, ByVal CodVersion As String, ByVal IdPeriodo As Integer, ByVal CodSeccion As String, ByVal TipoMonto As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr, rdr2 As OleDbDataReader
        Dim sql As String
        Dim i As Integer

        'Try
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "select D.CodDetalle, F.Rubro, D." & TipoMonto & " from DetalleS D join V_FlujoS F on D.CodVersion = F.CodVersion and D.IdPeriodo = F.IdPeriodo and D.CodDetalle = F.CodDetalle and D.Rubro = F.Rubro "
        sql &= "where F.CodVersion = '" & CodVersion & "' and F.IdPeriodo = " & IdPeriodo.ToString & " "
        If CodSeccion = "INGRES" Then
            sql &= "and F.CodSeccion = 'INGRES' "
        Else
            sql &= "and F.CodSeccion <> 'INGRES' "
        End If
        sql &= "order by Rubro"
        cmd = New OleDbCommand(sql, cn)
        rdr = cmd.ExecuteReader

        i = 0
        While rdr.Read
            hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")

            If Not FlagSemanal And (rdr("CodDetalle") = "VARIOS" Or rdr("CodDetalle") = "PRGNAC") Then
                i = i + 1
                sql = "select Rubro, " & TipoMonto & " from DetalleS "
                sql &= "where CodDetalle = '_" & rdr("CodDetalle").ToString.Substring(0, 5) & "' and IdPeriodo = " & IdPeriodo.ToString & " "
                sql &= "order by Rubro"
                cmd = New OleDbCommand(sql, cn)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    hoja.Cells(fil_ini + i, 2).InsertIndent(1)
                    hoja.Cells(fil_ini + i, 2).Font.Size = 10
                    hoja.Cells(fil_ini + i, 2).Value = rdr2("Rubro")
                    hoja.Cells(fil_ini + i, col).Value = rdr2(TipoMonto)
                    i = i + 1
                End While
            Else
                hoja.Cells(fil_ini + i, col).Value = rdr(TipoMonto)
                i = i + 1
            End If
        End While
        rdr.Close()

        If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraInversiones(ByVal hoja As Excel.Worksheet, ByVal TipoMonto As String, ByVal FechaIni As Date, ByVal FechaFin As Date, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col As Integer)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim rdr As OleDbDataReader
        Dim sql As String
        Dim i As Integer

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()

            If TipoMonto = "MontoBruto" Then TipoMonto = "MontoUSD" Else TipoMonto = "MontoBaseUSD"
            sql = "select Persona + ' - ' + Glosa as Rubro, sum(" & TipoMonto & ") as Monto from V_Movimiento M, Persona P "
            sql &= "where M.CodPersona = P.CodPersona and CodCuentaFC2 in (select CodCtaOrigen from CtaDetalleS where CodDetalle = 'INVERS') "
            sql &= "and Fecha between '" & FechaIni.ToString("yyyy-MM-dd") & "' and '" & FechaFin.ToString("yyyy-MM-dd") & "' "
            sql &= "group by Persona + ' - ' + Glosa order by 1"
            cmd = New OleDbCommand(sql, cn)
            rdr = cmd.ExecuteReader

            i = 0
            While rdr.Read
                hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")
                hoja.Cells(fil_ini + i, col).Value = rdr("Monto")
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

    Sub CreaIdMovimientoS(ByVal FlagNac As Boolean, ByVal FechaIni As String, ByVal FechaFin As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()
            If FlagNac Then sql = "sp_Crea_IdMovimientoS" Else sql = "sp_Crea_IdMovimientoS_Ext"
            cmd = New OleDbCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("FechaIni", FechaIni)
            cmd.Parameters.AddWithValue("FechaFin", FechaFin)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraDetalleCompleto(ByVal hoja As Excel.Worksheet, ByVal FlagNac As Boolean, ByVal FechaIni As String, ByVal FechaFin As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        Dim sql As String
        Dim i, fil_ini, fil As Integer
        Dim tabla As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            If FlagNac Then tabla = "V_Movimiento" Else tabla = "V_Movimiento_Ext"

            sql = "select distinct 'INGRES' as Seccion, IdCtaFlujo, Rubro " & _
                    "from " & tabla & " M, IdMovimientoS I, CtaFlujoS F " & _
                    "where M.IdMovimiento = I.IdMovimiento and I.CodDetalle = F.CodDetalle and CodSeccion = 'INGRES' " & _
                    "union select distinct 'EGRESO' as Seccion, IdCtaFlujo, Rubro " & _
                    "from " & tabla & " M, IdMovimientoS I, CtaFlujoS F " & _
                    "where M.IdMovimiento = I.IdMovimiento and I.CodDetalle = F.CodDetalle and CodSeccion <> 'INGRES' " & _
                    "order by 1 desc, 3"
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader

            fil_ini = 4
            i = 0
            While rdr.Read
                hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")
                hoja.Cells(fil_ini + i, 2).Font.Bold = True
                hoja.Cells(fil_ini + i, 2).Interior.Color = Excel.XlRgbColor.rgbYellow
                i = i + 1
                'hoja.Cells(fil_ini + i, 2).Value = "Persona"
                'hoja.Cells(fil_ini + i, 3).Value = "Area"
                'hoja.Cells(fil_ini + i, 4).Value = "Banco"
                'hoja.Cells(fil_ini + i, 5).Value = "Voucher"
                'hoja.Cells(fil_ini + i, 6).Value = "Item"
                'hoja.Cells(fil_ini + i, 7).Value = "Fecha"
                'hoja.Cells(fil_ini + i, 8).Value = "Glosa"
                'hoja.Cells(fil_ini + i, 9).Value = "Monto Neto"
                'hoja.Cells(fil_ini + i, 10).Value = "IGV"
                'hoja.Cells(fil_ini + i, 11).Value = "Monto Bruto"
                'hoja.Cells(fil_ini + i, 12).Value = "Cuenta FC"
                'hoja.Cells(fil_ini + i, 13).Value = "Grupo Movimiento"
                hoja.Range(hoja.Cells(fil_ini + i, 2), hoja.Cells(fil_ini + i, 13)).Font.Bold = True
                'i = i + 1
                fil = fil_ini + i + 1
                sql = "select Persona, A.Area, CodCuentaBanco as Banco, NroVoucher, NroItem, Fecha, Glosa, " & _
                        "I.Signo * MontoBaseUSD as MontoNeto, I.Signo * MontoIGVUSD as IGV, I.Signo * MontoUSD as MontoBruto, " & _
                        "Cuenta2 + ' [' + convert(varchar(3), CodCuenta2) + ']' as CuentaFC , GrupoMov + ' [' + convert(varchar(3), CodGrupoMov) + ']' as GrupoMovimiento " & _
                        "from " & tabla & " M, IdMovimientoS I, CtaFlujoS F, Persona P, Area A, V_CuentaFC C " & _
                        "where M.IdMovimiento = I.IdMovimiento And I.CodDetalle = F.CodDetalle " & _
                        "and M.CodPersona = P.CodPersona and M.CodArea2 = A.CodArea and M.CodCuentaFC2 = C.CodCuenta2 " & _
                        "and F.IdCtaFlujo = " & rdr("IdCtaFlujo").ToString & " and I.Rubro = '" & rdr("Rubro").ToString & "' " & _
                        "order by 1, 2, 3, 4, 5, 6, 7, 10"
                cmd = New SqlCommand(sql, cn)
                dtadap = New SqlDataAdapter(cmd)
                dtaset = New DataSet()
                dtadap.Fill(dtaset)
                Dim array As Object(,)
                array = DataSet2Array(dtaset.Tables(0), True)
                hoja.Range("B" & (fil - 1).ToString & ":M" & (fil_ini + i + dtaset.Tables(0).Rows.Count), Type.Missing).Value2 = array
                i = i + dtaset.Tables(0).Rows.Count + 1

                'rdr2 = cmd.ExecuteReader
                'While rdr2.Read
                '    hoja.Cells(fil_ini + i, 2).Value = rdr2("Persona")
                '    hoja.Cells(fil_ini + i, 3).Value = rdr2("Area")
                '    hoja.Cells(fil_ini + i, 4).Value = rdr2("CodCuentaBanco")
                '    hoja.Cells(fil_ini + i, 5).Value = rdr2("NroVoucher")
                '    hoja.Cells(fil_ini + i, 6).Value = rdr2("NroItem")
                '    hoja.Cells(fil_ini + i, 7).Value = Convert.ToDateTime(rdr2("Fecha")).ToString("yyyy-MM-dd")
                '    hoja.Cells(fil_ini + i, 8).Value = rdr2("Glosa")
                '    hoja.Cells(fil_ini + i, 9).Value = rdr2("MontoNeto")
                '    hoja.Cells(fil_ini + i, 10).Value = rdr2("IGV")
                '    hoja.Cells(fil_ini + i, 11).Value = rdr2("MontoBruto")
                '    hoja.Cells(fil_ini + i, 12).Value = rdr2("Cuenta2") & " [" & rdr2("CodCuenta2") & "]"
                '    hoja.Cells(fil_ini + i, 13).Value = rdr2("GrupoMov") & " [" & rdr2("CodGrupoMov") & "]"
                '    i = i + 1
                'End While
                hoja.Cells(fil_ini + i, 8).Value = "Total"
                hoja.Cells(fil_ini + i, 9).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                hoja.Cells(fil_ini + i, 10).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                hoja.Cells(fil_ini + i, 11).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                hoja.Range(hoja.Cells(fil_ini + i, 2), hoja.Cells(fil_ini + i, 11)).Font.Bold = True
                i = i + 2
            End While
            rdr.Close()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Sub GeneraFCDiario(ByVal excel As Excel.Application, ByVal reporte As Excel.Workbook, ByVal fil_ini As Integer, ByVal FlagNac As Boolean, ByVal TipoMonto As String, ByVal Dia As Date)
        Dim hoja As Excel.Worksheet
        Dim rango As Excel.Range
        Dim IdPeriodoAnt, IdPeriodoAct As Integer

        'Try
        hoja = reporte.Worksheets("Flujo ATV")
        rango = hoja.Rows("5:80")
        hoja = reporte.Worksheets("Flujo ATV Diario")
        rango.Copy(hoja.Rows((fil_ini + 1).ToString))
        'excel.Selection.Copy()
        'hoja = reporte.Worksheets("Flujo ATV Diario")
        'hoja.Rows(fil_ini.ToString).Select()
        'excel.Selection.Insert()

        hoja.Cells(fil_ini + 1, 2).Value = "DEL " & Dia.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        If TipoMonto = "MontoBruto" Then
            If Dia.DayOfWeek() <> DayOfWeek.Monday Then
                hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & Dia.AddDays(-1).ToString("dd.MM.yy") & " US$"
                hoja.Cells(fil_ini + 5, 3).Value = BuscaCierreCaja(Dia.AddDays(-1).ToString("yyyy-MM-dd"))
            Else
                hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & Dia.AddDays(-3).ToString("dd.MM.yy") & " US$"
                hoja.Cells(fil_ini + 5, 3).Value = BuscaCierreCaja(Dia.AddDays(-3).ToString("yyyy-MM-dd"))
            End If
            hoja.Cells(fil_ini + 74, 2).Value = "SALDO AL " & Dia.ToString("dd.MM.yy") & " US$"
            hoja.Rows((fil_ini + 75).ToString & ":" & (fil_ini + 75).ToString).EntireRow.Hidden = True
        Else
            hoja.Cells(fil_ini, 2).Value = "FLUJO DE CAJA EJECUTADO RAG"
            If Dia.DayOfWeek() <> DayOfWeek.Monday Then
                hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & Dia.AddDays(-1).ToString("dd.MM.yy") & " US$"
            Else
                hoja.Cells(fil_ini + 5, 2).Value = "SALDO AL " & Dia.AddDays(-3).ToString("dd.MM.yy") & " US$"
            End If
            hoja.Cells(fil_ini + 74, 2).Value = "SALDO SIN IGV AL " & Dia.ToString("dd.MM.yy") & " US$"
            hoja.Cells(fil_ini + 75, 2).Value = "SALDO CON IGV AL " & Dia.ToString("dd.MM.yy") & " US$"
        End If
        IdPeriodoAnt = Convert.ToInt32(Dia.ToString("yyyyMM"))
        IdPeriodoAct = Convert.ToInt32(Dia.ToString("yyyyMM"))

        GeneraDetalleD(False, hoja, "E", Dia, "INGRES", TipoMonto, fil_ini + 6, fil_ini + 10, 3)
        GeneraDetalleD(False, hoja, "E", Dia, "EGRESO", TipoMonto, fil_ini + 13, fil_ini + 72, 3)

        'CreaDetalleS(FlagNac, IdPeriodoAnt, IdPeriodoAct, Dia, Dia)
        'GeneraDetalle2(False, hoja, "E", IdPeriodoAct, "INGRES", TipoMonto, fil_ini + 6, fil_ini + 10, 3)
        'GeneraDetalle2(False, hoja, "E", IdPeriodoAct, "EGRESO", TipoMonto, fil_ini + 13, fil_ini + 62, 3)

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub CreaDetalleD(ByVal FlagNac As Boolean, ByVal FechaIni As Date, ByVal FechaFin As Date)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            If FlagNac Then sql = "sp_Crea_DetalleD" Else sql = "sp_Crea_DetalleD_Ext"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("FechaIni", FechaIni.ToString("yyyy-MM-dd"))
            cmd.Parameters.AddWithValue("FechaFin", FechaFin.ToString("yyyy-MM-dd"))
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraDetalleD(ByVal FlagSemanal As Boolean, ByVal hoja As Excel.Worksheet, ByVal CodVersion As String, ByVal Dia As Date, ByVal CodSeccion As String, ByVal TipoMonto As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr, rdr2 As SqlDataReader
        Dim sql As String
        Dim i As Integer

        'Try
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodDetalle, Rubro, " & TipoMonto & " from DetalleD D "
        sql &= "where CodVersion = '" & CodVersion & "' and Fecha = '" & Dia.ToString & "' "
        If CodSeccion = "INGRES" Then
            sql &= "and CodSeccion = 'INGRES' "
        Else
            sql &= "and CodSeccion <> 'INGRES' and CodSeccion <> 'ZZZZZZ' "
        End If
        sql &= "order by Rubro"
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader

        i = 0
        While rdr.Read
            hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")

            If Not FlagSemanal And (rdr("CodDetalle") = "VARIOS" Or rdr("CodDetalle") = "PRGNAC") Then
                i = i + 1
                sql = "select Rubro, " & TipoMonto & " from DetalleD "
                sql &= "where CodDetalle = '_" & rdr("CodDetalle").ToString.Substring(0, 5) & "' and Fecha = '" & Dia.ToString & "' "
                sql &= "order by Rubro"
                cmd = New SqlCommand(sql, cn)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    hoja.Cells(fil_ini + i, 2).InsertIndent(1)
                    hoja.Cells(fil_ini + i, 2).Font.Size = 10
                    hoja.Cells(fil_ini + i, 2).Value = rdr2("Rubro")
                    hoja.Cells(fil_ini + i, col).Value = rdr2(TipoMonto)
                    i = i + 1
                End While
            Else
                hoja.Cells(fil_ini + i, col).Value = rdr(TipoMonto)
                i = i + 1
            End If
        End While
        rdr.Close()

        If fil_fin >= fil_ini + i Then hoja.Rows((fil_ini + i).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraDetalleDiarioCompleto(ByVal hoja As Excel.Worksheet, ByVal FlagNac As Boolean, ByVal FechaIni As String, ByVal FechaFin As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr, rdr0 As SqlDataReader
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim i, fil_ini, fil As Integer
        Dim tabla As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            If FlagNac Then tabla = "V_Movimiento" Else tabla = "V_Movimiento_Ext"

            sql = "select distinct Fecha from " & tabla & " " & _
                    "where Fecha >= '" & FechaIni & "' and Fecha <= '" & FechaFin & "' order by 1"
            cmd = New SqlCommand(sql, cn)
            rdr0 = cmd.ExecuteReader

            fil_ini = 4
            i = 0
            While rdr0.Read
                hoja.Cells(fil_ini + i, 2).Value = "Día " & Convert.ToDateTime(rdr0("Fecha")).ToString("dd/MM/yyyy")
                hoja.Cells(fil_ini + i, 2).Font.Bold = True
                hoja.Cells(fil_ini + i, 2).Interior.Color = Excel.XlRgbColor.rgbYellow
                i = i + 1
                'sql = "select distinct 'INGRES' as Seccion, IdCtaFlujo, I.CodDetalle, Rubro " & _
                '        "from " & tabla & " M, IdMovimientoS I, CtaFlujoS F " & _
                '        "where M.IdMovimiento = I.IdMovimiento and I.CodDetalle = F.CodDetalle and CodSeccion = 'INGRES' " & _
                '        "and Fecha = '" & Convert.ToDateTime(rdr0("Fecha")).ToString("yyyy-MM-dd") & "' " & _
                '        "union select distinct 'EGRESO' as Seccion, IdCtaFlujo, I.CodDetalle, Rubro " & _
                '        "from " & tabla & " M, IdMovimientoS I, CtaFlujoS F " & _
                '        "where M.IdMovimiento = I.IdMovimiento and I.CodDetalle = F.CodDetalle and CodSeccion <> 'INGRES' " & _
                '        "and Fecha = '" & Convert.ToDateTime(rdr0("Fecha")).ToString("yyyy-MM-dd") & "' " & _
                '        "order by 1 desc, 4"
                sql = "select case CodSeccion when 'INGRES' then 'INGRES' else 'EGRESO' end as CodSeccion, CodDetalle, Rubro " & _
                        "from DetalleD " & _
                        "where CodVersion = 'E' and Fecha = '" & rdr0("Fecha").ToString() & "' and CodSeccion <> 'ZZZZZZ'" & _
                        "order by 1 desc, 3"
                cmd = New SqlCommand(sql, cn)
                rdr = cmd.ExecuteReader

                While rdr.Read
                    hoja.Cells(fil_ini + i, 2).Value = rdr("Rubro")
                    hoja.Cells(fil_ini + i, 2).Font.Bold = True
                    hoja.Cells(fil_ini + i, 2).Interior.Color = Excel.XlRgbColor.rgbYellow
                    i = i + 1
                    'hoja.Cells(fil_ini + i, 2).Value = "Persona"
                    'hoja.Cells(fil_ini + i, 3).Value = "Area"
                    'hoja.Cells(fil_ini + i, 4).Value = "Banco"
                    'hoja.Cells(fil_ini + i, 5).Value = "Voucher"
                    'hoja.Cells(fil_ini + i, 6).Value = "Item"
                    'hoja.Cells(fil_ini + i, 7).Value = "Fecha"
                    'hoja.Cells(fil_ini + i, 8).Value = "Glosa"
                    'hoja.Cells(fil_ini + i, 9).Value = "Monto Neto"
                    'hoja.Cells(fil_ini + i, 10).Value = "IGV"
                    'hoja.Cells(fil_ini + i, 11).Value = "Monto Bruto"
                    'hoja.Cells(fil_ini + i, 12).Value = "Cuenta FC"
                    'hoja.Cells(fil_ini + i, 13).Value = "Grupo Movimiento"
                    hoja.Range(hoja.Cells(fil_ini + i, 2), hoja.Cells(fil_ini + i, 13)).Font.Bold = True
                    'i = i + 1
                    fil = fil_ini + i + 1
                    sql = "select Persona, A.Area, CodCuentaBanco as Banco, NroVoucher as Voucher, NroItem as Item, Fecha, Glosa, " & _
                            "I.Signo * MontoBaseUSD as MontoNeto, I.Signo * MontoIGVUSD as IGV, I.Signo * MontoUSD as MontoBruto, " & _
                            "Cuenta2 + ' [' + convert(varchar(3), CodCuenta2) + ']' as CuentaFC , GrupoMov + ' [' + convert(varchar(3), CodGrupoMov) + ']' as GrupoMovimiento " & _
                            "from " & tabla & " M, IdMovimientoS I, Persona P, Area A, V_CuentaFC C " & _
                            "where M.IdMovimiento = I.IdMovimiento " & _
                            "and M.CodPersona = P.CodPersona and M.CodArea2 = A.CodArea and M.CodCuentaFC2 = C.CodCuenta2 " & _
                            "and I.CodDetalle = '" & rdr("CodDetalle").ToString & "' and I.Rubro = '" & rdr("Rubro").ToString & "' " & _
                            "and Fecha = '" & Convert.ToDateTime(rdr0("Fecha")).ToString("yyyy-MM-dd") & "' "
                    If rdr("CodDetalle") = "PRGNAC" Or rdr("CodDetalle") = "VARIOS" Then
                        sql &= "order by 2, 1, 3, 4, 5, 6, 7, 10"
                    Else
                        sql &= "order by 1, 2, 3, 4, 5, 6, 7, 10"
                    End If
                    cmd = New SqlCommand(sql, cn)
                    dtadap = New SqlDataAdapter(cmd)
                    dtaset = New DataSet()
                    dtadap.Fill(dtaset)
                    Dim array As Object(,)
                    array = DataSet2Array(dtaset.Tables(0), True)
                    hoja.Range("B" & (fil - 1).ToString & ":M" & (fil_ini + i + dtaset.Tables(0).Rows.Count), Type.Missing).Value2 = array
                    i = i + dtaset.Tables(0).Rows.Count + 1
                    'rdr2 = cmd.ExecuteReader
                    'While rdr2.Read
                    '    hoja.Cells(fil_ini + i, 2).Value = rdr2("Persona")
                    '    hoja.Cells(fil_ini + i, 3).Value = rdr2("Area")
                    '    hoja.Cells(fil_ini + i, 4).Value = rdr2("CodCuentaBanco")
                    '    hoja.Cells(fil_ini + i, 5).Value = rdr2("NroVoucher")
                    '    hoja.Cells(fil_ini + i, 6).Value = rdr2("NroItem")
                    '    hoja.Cells(fil_ini + i, 7).Value = Convert.ToDateTime(rdr2("Fecha")).ToString("yyyy-MM-dd")
                    '    hoja.Cells(fil_ini + i, 8).Value = rdr2("Glosa")
                    '    hoja.Cells(fil_ini + i, 9).Value = rdr2("MontoNeto")
                    '    hoja.Cells(fil_ini + i, 10).Value = rdr2("IGV")
                    '    hoja.Cells(fil_ini + i, 11).Value = rdr2("MontoBruto")
                    '    hoja.Cells(fil_ini + i, 12).Value = rdr2("Cuenta2") & " [" & rdr2("CodCuenta2") & "]"
                    '    hoja.Cells(fil_ini + i, 13).Value = rdr2("GrupoMov") & " [" & rdr2("CodGrupoMov") & "]"
                    '    i = i + 1
                    'End While
                    hoja.Cells(fil_ini + i, 8).Value = "Total"
                    hoja.Cells(fil_ini + i, 9).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                    hoja.Cells(fil_ini + i, 10).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                    hoja.Cells(fil_ini + i, 11).FormulaR1C1 = "=SUM(R" & fil.ToString & "C[0]:R" & (fil_ini + i - 1).ToString & "C[0])"
                    hoja.Range(hoja.Cells(fil_ini + i, 2), hoja.Cells(fil_ini + i, 11)).Font.Bold = True
                    i = i + 2
                End While
                rdr.Close()
                'i = i + 2
            End While
            rdr0.Close()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    'Function DataSet2Array(ByVal dt As DataTable) As Object(,)
    '    ' Copy the DataTable to an object array
    '    Dim rawData(dt.Rows.Count, dt.Columns.Count - 1) As Object

    '    'Copy the column names to the first row of the object array
    '    For col = 0 To dt.Columns.Count - 1
    '        rawData(0, col) = dt.Columns(col).ColumnName
    '    Next

    '    ' Copy the values to the object array
    '    For col = 0 To dt.Columns.Count - 1
    '        For row = 0 To dt.Rows.Count - 1
    '            If Not IsDBNull(dt.Rows(row).ItemArray(col)) Then
    '                rawData(row + 1, col) = dt.Rows(row).ItemArray(col)
    '            End If
    '        Next
    '    Next

    '    Return rawData
    'End Function

End Module

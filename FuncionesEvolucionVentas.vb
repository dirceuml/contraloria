Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesEvolucionVentas

    Private Function BuscaVenta(ByVal CodMedio As String, ByVal FechaIni As String, ByVal FechaFin As String, ByVal FlagPublicidad As Boolean, ByVal FlagContado As Boolean, ByVal FlagATV As Boolean) As Double
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim Venta As Double = 0
        Dim CodRubro As String = ""

        If FlagPublicidad And FlagContado Then
            CodRubro = "1PUBLI"
        ElseIf FlagPublicidad And Not FlagContado Then
            CodRubro = "2CANJE"
        ElseIf Not FlagPublicidad And FlagContado Then
            CodRubro = "3OTROS"
        End If

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(sum(VentaUSD), 0) as VentaUSD " & _
                "from Venta V, VentaRubro R where V.CodMedio = R.CodMedio and V.CodFormaPago = R.CodFormaPago and V.Flag = R.Flag " & _
                "and V.CodMedio = " & CodMedio & " and FechaEmision between '" & FechaIni & "' and '" & FechaFin & "' " & _
                "and CodRubro = '" & CodRubro & "'"
        If FlagATV Then sql += "and CodCliente = 613 "

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then Venta = Convert.ToDouble(rdr("VentaUSD"))
        rdr.Close()
        cn.Close()

        cn.Close()

        Return Venta
    End Function

    'Private Function BuscaVenta(ByVal CodMedio As String, ByVal FechaIni As String, ByVal FechaFin As String, ByVal FlagPublicidad As Boolean, ByVal FlagContado As Boolean, ByVal FlagATV As Boolean) As Double
    '    Dim sql As String
    '    Dim cn As SqlConnection
    '    Dim cmd As SqlCommand
    '    Dim rdr As SqlDataReader
    '    Dim Venta As Double = 0

    '    If CodMedio <> "2" And CodMedio <> "11" Then
    '        cn = New SqlConnection
    '        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
    '        cn.Open()

    '        sql = "select isnull(sum(VentaUSD), 0) as VentaUSD " & _
    '                "from Venta where CodMedio = " & CodMedio & " and FechaEmision between '" & FechaIni & "' and '" & FechaFin & "' "
    '        If FlagPublicidad Then sql += "and Flag = 1 " Else sql += "and Flag = 0 "
    '        If FlagContado Then sql += "and CodFormaPago not in (4, 5, 7) " Else sql += "and CodFormaPago in (4, 5, 7) "
    '        If FlagATV Then sql += "and CodCliente = 613 "

    '        cmd = New SqlCommand(sql, cn)
    '        rdr = cmd.ExecuteReader
    '        If rdr.Read() Then Venta = Convert.ToDouble(rdr("VentaUSD"))
    '        rdr.Close()
    '        cn.Close()

    '        cn.Close()
    '    Else
    '        Dim CodRubro As String = ""
    '        If FlagPublicidad And FlagContado Then
    '            CodRubro = "1PUBLI"
    '        ElseIf FlagPublicidad And Not FlagContado Then
    '            CodRubro = "2CANJE"
    '        ElseIf Not FlagPublicidad And FlagContado Then
    '            CodRubro = "3OTROS"
    '        End If
    '        If CodRubro <> "" Then Venta = BuscaVenta2(CodMedio, FechaIni, FechaFin, CodRubro, FlagATV)
    '    End If

    '    Return Venta
    'End Function

    Private Function BuscaVentaNeta(ByVal CodMedio As String, ByVal FechaIni As String, ByVal FechaFin As String, ByVal FlagPublicidad As Boolean, ByVal FlagContado As String, ByVal FlagATV As Boolean) As Double
        Dim VentaNeta As Double = 0
        Dim VentaNeta2 As Double = 0

        If CodMedio = "2" Or CodMedio = "3" Or CodMedio = "6" Then
            VentaNeta = BuscaVenta(CodMedio, FechaIni, FechaFin, FlagPublicidad, FlagContado, FlagATV)
        ElseIf CodMedio = "1" Then
            VentaNeta = BuscaVenta("1", FechaIni, FechaFin, FlagPublicidad, FlagContado, False) - _
                        BuscaVenta("2", FechaIni, FechaFin, FlagPublicidad, FlagContado, True) - _
                        BuscaVenta("3", FechaIni, FechaFin, FlagPublicidad, FlagContado, True) - _
                        BuscaVenta("6", FechaIni, FechaFin, FlagPublicidad, FlagContado, True)
        ElseIf CodMedio = "0" Then
            VentaNeta = BuscaVenta("1", FechaIni, FechaFin, FlagPublicidad, FlagContado, False) - _
                        BuscaVenta("2", FechaIni, FechaFin, FlagPublicidad, FlagContado, True) - _
                        BuscaVenta("3", FechaIni, FechaFin, FlagPublicidad, FlagContado, True) - _
                        BuscaVenta("6", FechaIni, FechaFin, FlagPublicidad, FlagContado, True) + _
                        BuscaVenta("2", FechaIni, FechaFin, FlagPublicidad, FlagContado, False) + _
                        BuscaVenta("3", FechaIni, FechaFin, FlagPublicidad, FlagContado, False) + _
                        BuscaVenta("6", FechaIni, FechaFin, FlagPublicidad, FlagContado, False)
        End If

        Return VentaNeta
    End Function

    Private Function BuscaVentaPpto(ByVal IdPeriodo As String, ByVal CodMedio As String, ByVal CodRubro As String, ByVal FlagAcumulado As Boolean) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim VentaPpto As Double = 0

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(sum(VentaUSD), 0) as VentaUSD from VentaPpto " & _
                "where CodRubro = '" & CodRubro & "' "
        If CodMedio <> "0" Then sql &= "and CodMedio = " & CodMedio & " "
        If Not FlagAcumulado Then sql &= "and IdPeriodo = " & IdPeriodo & " " Else sql &= "and IdPeriodo between " & IdPeriodo.Substring(0, 4) & "01 and " & IdPeriodo & " "

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then VentaPpto = Convert.ToDouble(rdr("VentaUSD"))
        rdr.Close()
        cn.Close()

        cn.Close()

        Return VentaPpto
    End Function

    'Private Function BuscaVenta2(ByVal CodMedio As String, ByVal FechaIni As String, ByVal FechaFin As String, ByVal CodRubro As String, ByVal FlagATV As Boolean) As Double
    '    Dim sql As String
    '    Dim cn As New SqlConnection
    '    Dim cmd As SqlCommand
    '    Dim rdr As SqlDataReader
    '    Dim Venta2 As Double = 0

    '    cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
    '    cn.Open()

    '    sql = "select isnull(sum(VentaAcumUSD), 0) as VentaAcumUSD from Venta2 V, "
    '    sql &= "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, Medio, CodRubro, Cliente, max(Fecha) as Fecha " & _
    '            "from Venta2 " & _
    '            "where Fecha between '" & FechaIni & "' and '" & FechaFin & "' and CodRubro = '" & CodRubro & "' "
    '    If CodMedio = "2" Then
    '        sql &= "and Medio = 'Global' "
    '    ElseIf CodMedio = "11" Then
    '        sql &= "and Medio = 'ATV_SUR' "
    '    Else
    '        sql &= "and 1 = 0 "
    '    End If
    '    If FlagATV Then sql &= "and Cliente = 'ATV' " Else sql &= "and Cliente = 'Total' "
    '    sql = sql & "group by year(Fecha) * 100 + month(Fecha), Medio, CodRubro, Cliente) T " & _
    '            "where V.Medio = T.Medio and V.Fecha = T.Fecha and V.CodRubro = T.CodRubro and V.Cliente = T.Cliente "

    '    cmd = New SqlCommand(sql, cn)
    '    rdr = cmd.ExecuteReader
    '    If rdr.Read() Then Venta2 = Convert.ToDouble(rdr("VentaAcumUSD"))
    '    rdr.Close()
    '    cn.Close()

    '    cn.Close()

    '    Return Venta2
    'End Function

    Sub GeneraReportesEvolucionVentas(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal Fecha As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("PERU")
        GeneraEvolucionVentas(hoja, "0", Convert.ToDateTime(Fecha))

        hoja = reporte.Worksheets("ATV")
        GeneraEvolucionVentas(hoja, "1", Convert.ToDateTime(Fecha))

        hoja = reporte.Worksheets("GLOBAL")
        GeneraEvolucionVentas(hoja, "2", Convert.ToDateTime(Fecha))

        hoja = reporte.Worksheets("LA TELE")
        GeneraEvolucionVentas(hoja, "3", Convert.ToDateTime(Fecha))

        hoja = reporte.Worksheets("ATV_SUR")
        GeneraEvolucionVentas(hoja, "6", Convert.ToDateTime(Fecha))

        hoja = reporte.Worksheets("Anexo ATV Fact Compartida")
        GeneraAnexoFacturacionCompartida(hoja, Convert.ToDateTime(Fecha))

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

    Private Sub GeneraEvolucionVentas(ByVal hoja As Excel.Worksheet, ByVal CodMedio As String, ByVal Fecha As Date)
        Dim FechaIni As String, FechaFin As String
        Dim fil_ini As Integer

        If Fecha.ToString("yyyyMM") <> Fecha.AddDays(1).ToString("yyyyMM") Then
            hoja.Cells(4, 2).Value = "INFORME MENSUAL " & hoja.Name & " - EVOLUCION DE VENTAS"
            hoja.Cells(9, 4).Value = Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(9, 8).Value = "ENERO - " & hoja.Cells(9, 4).Value
        Else
            hoja.Cells(4, 2).Value = "INFORME SEMANAL " & hoja.Name & " - EVOLUCION DE VENTAS"
            hoja.Cells(9, 4).Value = "DEL 01 " & Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper & _
                                        " AL " & Fecha.ToString("dd MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(9, 8).Value = "DEL 01 ENERO AL " & Fecha.ToString("dd MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        End If

        hoja.Cells(9, 5).Value = Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(19, 4).Value = hoja.Cells(9, 4).Value
        hoja.Cells(19, 5).Value = hoja.Cells(9, 5).Value 'hoja.Cells(9, 4).Value

        hoja.Cells(9, 9).Value = "ENERO - " & hoja.Cells(9, 5).Value
        hoja.Cells(19, 8).Value = hoja.Cells(9, 8).Value
        hoja.Cells(19, 9).Value = hoja.Cells(9, 9).Value

        fil_ini = 10

        FechaIni = Fecha.ToString("yyyy-MM") & "-01"
        FechaFin = Fecha.ToString("yyyy-MM-dd")

        hoja.Cells(6, 10).Value = FechaFin

        hoja.Cells(fil_ini, 4).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini + 1, 4).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 2, 4).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, True, False) + BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, False, False)

        hoja.Cells(fil_ini + 10, 4).Value = hoja.Cells(fil_ini, 4).Value
        hoja.Cells(fil_ini + 11, 4).Value = hoja.Cells(fil_ini + 1, 4).Value
        hoja.Cells(fil_ini + 12, 4).Value = hoja.Cells(fil_ini + 2, 4).Value

        'Presupuesto
        hoja.Cells(fil_ini, 5).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "1PUBLI", False)
        hoja.Cells(fil_ini + 1, 5).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "2CANJE", False)
        hoja.Cells(fil_ini + 2, 5).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "3OTROS", False)

        'Acumulado
        FechaIni = Fecha.ToString("yyyy") & "-01-01"
        FechaFin = Fecha.ToString("yyyy-MM-dd")

        hoja.Cells(fil_ini, 8).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini + 1, 8).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 2, 8).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, True, False) + BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, False, False)

        hoja.Cells(fil_ini + 10, 8).Value = hoja.Cells(fil_ini, 8).Value
        hoja.Cells(fil_ini + 11, 8).Value = hoja.Cells(fil_ini + 1, 8).Value
        hoja.Cells(fil_ini + 12, 8).Value = hoja.Cells(fil_ini + 2, 8).Value

        'Presupuesto Acumulado
        hoja.Cells(fil_ini, 9).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "1PUBLI", True)
        hoja.Cells(fil_ini + 1, 9).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "2CANJE", True)
        hoja.Cells(fil_ini + 2, 9).Value = BuscaVentaPpto(Fecha.ToString("yyyyMM"), CodMedio, "3OTROS", True)

        'Año Anterior
        FechaIni = (Year(Fecha) - 1).ToString & Fecha.ToString("-MM") & "-01"
        FechaFin = Convert.ToDateTime(FechaIni).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        hoja.Cells(fil_ini + 10, 5).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini + 11, 5).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 12, 5).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, True, False) + BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, False, False)

        'Acumulado Año Anterior
        FechaIni = (Year(Fecha) - 1).ToString & "-01-01"
        FechaFin = Convert.ToDateTime((Year(Fecha) - 1).ToString & "-" & Fecha.ToString("MM") & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        hoja.Cells(fil_ini + 10, 9).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini + 11, 9).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 12, 9).Value = BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, True, False) + BuscaVentaNeta(CodMedio, FechaIni, FechaFin, False, False, False)

    End Sub

    Private Sub GeneraAnexoFacturacionCompartida(ByVal hoja As Excel.Worksheet, ByVal Fecha As Date)
        Dim FechaIni As String, FechaFin As String
        Dim fil_ini As Integer

        fil_ini = 10

        If Fecha.ToString("yyyyMM") <> Fecha.AddDays(1).ToString("yyyyMM") Then
            hoja.Cells(4, 3).Value = "INFORME MENSUAL ATV - EVOLUCION DE VENTAS  NETAS  DE"
            hoja.Cells(9, 5).Value = Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(9, 11).Value = "ENERO - " & hoja.Cells(9, 5).Value
        Else
            hoja.Cells(4, 3).Value = "INFORME SEMANAL ATV - EVOLUCION DE VENTAS  NETAS  DE"
            hoja.Cells(9, 5).Value = "DEL 01 " & Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper & _
                                        " AL " & Fecha.ToString("dd MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            hoja.Cells(9, 11).Value = "DEL 01 ENERO AL " & Fecha.ToString("dd MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        End If

        hoja.Cells(9, 6).Value = hoja.Cells(9, 5).Value
        hoja.Cells(9, 7).Value = hoja.Cells(9, 5).Value
        hoja.Cells(9, 8).Value = hoja.Cells(9, 5).Value
        hoja.Cells(9, 9).Value = hoja.Cells(9, 5).Value
        hoja.Cells(18, 5).Value = Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(18, 6).Value = hoja.Cells(18, 5).Value
        hoja.Cells(18, 7).Value = hoja.Cells(18, 5).Value
        hoja.Cells(18, 8).Value = hoja.Cells(18, 5).Value
        hoja.Cells(18, 9).Value = hoja.Cells(18, 5).Value

        hoja.Cells(9, 12).Value = hoja.Cells(9, 11).Value
        hoja.Cells(9, 13).Value = hoja.Cells(9, 11).Value
        hoja.Cells(9, 14).Value = hoja.Cells(9, 11).Value
        hoja.Cells(9, 15).Value = hoja.Cells(9, 11).Value
        hoja.Cells(18, 11).Value = "ENERO - " & Fecha.ToString("MMMM", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(18, 12).Value = hoja.Cells(18, 11).Value
        hoja.Cells(18, 13).Value = hoja.Cells(18, 11).Value
        hoja.Cells(18, 14).Value = hoja.Cells(18, 11).Value
        hoja.Cells(18, 15).Value = hoja.Cells(18, 11).Value

        FechaIni = Fecha.ToString("yyyy-MM") & "-01"
        FechaFin = Fecha.ToString("yyyy-MM-dd")

        hoja.Cells(6, 15).Value = FechaFin

        hoja.Cells(fil_ini, 6).Value = BuscaVenta("1", FechaIni, FechaFin, True, True, False) 'ATV Publicidad Contado
        hoja.Cells(fil_ini + 1, 6).Value = BuscaVenta("1", FechaIni, FechaFin, True, False, False) 'ATV Publicidad Canje
        hoja.Cells(fil_ini + 2, 6).Value = BuscaVenta("1", FechaIni, FechaFin, False, True, False) + BuscaVenta("1", FechaIni, FechaFin, False, False, False) 'ATV Otros Contado + Canje

        hoja.Cells(fil_ini, 7).Value = BuscaVenta("2", FechaIni, FechaFin, True, True, True) 'Global Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 7).Value = BuscaVenta("2", FechaIni, FechaFin, True, False, True) 'Global Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 7).Value = BuscaVenta("2", FechaIni, FechaFin, False, True, True) 'Global Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 8).Value = BuscaVenta("3", FechaIni, FechaFin, True, True, True) 'La Tele Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 8).Value = BuscaVenta("3", FechaIni, FechaFin, True, False, True) 'La Tele Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 8).Value = BuscaVenta("3", FechaIni, FechaFin, False, True, True) + BuscaVenta("3", FechaIni, FechaFin, False, False, True) 'La Tele Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 9).Value = BuscaVenta("6", FechaIni, FechaFin, True, True, True) 'ATV_SUR Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 9).Value = BuscaVenta("6", FechaIni, FechaFin, True, False, True) 'ATV_SUR Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 9).Value = BuscaVenta("6", FechaIni, FechaFin, False, True, True) + BuscaVenta("6", FechaIni, FechaFin, False, False, True) 'ATV_SUR Otros Contado + Canje (Por ATV)

        'Acumulado
        FechaIni = Fecha.ToString("yyyy") & "-01-01"
        FechaFin = Fecha.ToString("yyyy-MM-dd")

        hoja.Cells(fil_ini, 12).Value = BuscaVenta("1", FechaIni, FechaFin, True, True, False) 'ATV Publicidad Contado
        hoja.Cells(fil_ini + 1, 12).Value = BuscaVenta("1", FechaIni, FechaFin, True, False, False) 'ATV Publicidad Canje
        hoja.Cells(fil_ini + 2, 12).Value = BuscaVenta("1", FechaIni, FechaFin, False, True, False) + BuscaVenta("1", FechaIni, FechaFin, False, False, False) 'ATV Otros Contado + Canje

        hoja.Cells(fil_ini, 13).Value = BuscaVenta("2", FechaIni, FechaFin, True, True, True) 'Global Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 13).Value = BuscaVenta("2", FechaIni, FechaFin, True, False, True) 'Global Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 13).Value = BuscaVenta("2", FechaIni, FechaFin, False, True, True) 'Global Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 14).Value = BuscaVenta("3", FechaIni, FechaFin, True, True, True) 'La Tele Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 14).Value = BuscaVenta("3", FechaIni, FechaFin, True, False, True) 'La Tele Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 14).Value = BuscaVenta("3", FechaIni, FechaFin, False, True, True) + BuscaVenta("3", FechaIni, FechaFin, False, False, True) 'La Tele Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 15).Value = BuscaVenta("6", FechaIni, FechaFin, True, True, True) 'ATV_SUR Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 15).Value = BuscaVenta("6", FechaIni, FechaFin, True, False, True) 'ATV_SUR Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 15).Value = BuscaVenta("6", FechaIni, FechaFin, False, True, True) + BuscaVenta("6", FechaIni, FechaFin, False, False, True) 'ATV_SUR Otros Contado + Canje (Por ATV)

        'Año Anterior
        FechaIni = (Year(Fecha) - 1).ToString & Fecha.ToString("-MM") & "-01"
        FechaFin = Convert.ToDateTime(FechaIni).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        fil_ini = 19

        hoja.Cells(fil_ini, 6).Value = BuscaVenta("1", FechaIni, FechaFin, True, True, False) 'ATV Publicidad Contado
        hoja.Cells(fil_ini + 1, 6).Value = BuscaVenta("1", FechaIni, FechaFin, True, False, False) 'ATV Publicidad Canje
        hoja.Cells(fil_ini + 2, 6).Value = BuscaVenta("1", FechaIni, FechaFin, False, True, False) + BuscaVenta("1", FechaIni, FechaFin, False, False, False) 'ATV Otros Contado + Canje

        hoja.Cells(fil_ini, 7).Value = BuscaVenta("2", FechaIni, FechaFin, True, True, True) 'Global Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 7).Value = BuscaVenta("2", FechaIni, FechaFin, True, False, True) 'Global Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 7).Value = BuscaVenta("2", FechaIni, FechaFin, False, True, True) 'Global Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 8).Value = BuscaVenta("3", FechaIni, FechaFin, True, True, True) 'La Tele Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 8).Value = BuscaVenta("3", FechaIni, FechaFin, True, False, True) 'La Tele Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 8).Value = BuscaVenta("3", FechaIni, FechaFin, False, True, True) + BuscaVenta("3", FechaIni, FechaFin, False, False, True) 'La Tele Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 9).Value = BuscaVenta("6", FechaIni, FechaFin, True, True, True) 'ATV_SUR Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 9).Value = BuscaVenta("6", FechaIni, FechaFin, True, False, True) 'ATV_SUR Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 9).Value = BuscaVenta("6", FechaIni, FechaFin, False, True, True) + BuscaVenta("6", FechaIni, FechaFin, False, False, True) 'ATV_SUR Otros Contado + Canje (Por ATV)

        'Acumulado Año Anterior
        FechaIni = (Year(Fecha) - 1).ToString & "-01-01"
        FechaFin = Convert.ToDateTime((Year(Fecha) - 1).ToString & "-" & Fecha.ToString("MM") & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        hoja.Cells(fil_ini, 12).Value = BuscaVenta("1", FechaIni, FechaFin, True, True, False) 'ATV Publicidad Contado
        hoja.Cells(fil_ini + 1, 12).Value = BuscaVenta("1", FechaIni, FechaFin, True, False, False) 'ATV Publicidad Canje
        hoja.Cells(fil_ini + 2, 12).Value = BuscaVenta("1", FechaIni, FechaFin, False, True, False) + BuscaVenta("1", FechaIni, FechaFin, False, False, False) 'ATV Otros Contado + Canje

        hoja.Cells(fil_ini, 13).Value = BuscaVenta("2", FechaIni, FechaFin, True, True, True) 'Global Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 13).Value = BuscaVenta("2", FechaIni, FechaFin, True, False, True) 'Global Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 13).Value = BuscaVenta("2", FechaIni, FechaFin, False, True, True) 'Global Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 14).Value = BuscaVenta("3", FechaIni, FechaFin, True, True, True) 'La Tele Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 14).Value = BuscaVenta("3", FechaIni, FechaFin, True, False, True) 'La Tele Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 14).Value = BuscaVenta("3", FechaIni, FechaFin, False, True, True) + BuscaVenta("3", FechaIni, FechaFin, False, False, True) 'La Tele Otros Contado + Canje (Por ATV)

        hoja.Cells(fil_ini, 15).Value = BuscaVenta("6", FechaIni, FechaFin, True, True, True) 'ATV_SUR Publicidad Contado (Por ATV)
        hoja.Cells(fil_ini + 1, 15).Value = BuscaVenta("6", FechaIni, FechaFin, True, False, True) 'ATV_SUR Publicidad Canje (Por ATV)
        hoja.Cells(fil_ini + 2, 15).Value = BuscaVenta("6", FechaIni, FechaFin, False, True, True) + BuscaVenta("6", FechaIni, FechaFin, False, False, True) 'ATV_SUR Otros Contado + Canje (Por ATV)

    End Sub

    Sub GeneraReportesVentasGrupo(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        Dim IdPeriodoAnt As String
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("VENTAS ACT")
        hoja.Name = "VENTAS " & IdPeriodo.Substring(0, 4)
        GeneraVentasGrupo(hoja, IdPeriodo)

        IdPeriodoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & IdPeriodo.Substring(4, 2)
        hoja = reporte.Worksheets("VENTAS ANT")
        hoja.Name = "VENTAS " & IdPeriodoAnt.Substring(0, 4)
        GeneraVentasGrupo(hoja, IdPeriodoAnt)

        hoja = reporte.Worksheets("COMPARAT. ACT - ANT")
        hoja.Name = "COMPARAT. " & IdPeriodo.Substring(0, 4) & " - " & IdPeriodoAnt.Substring(0, 4)
        hoja.Cells(8, 3).Value = IdPeriodo.Substring(0, 4)
        hoja.Cells(8, 5).Value = IdPeriodo.Substring(0, 4)
        hoja.Cells(8, 7).Value = IdPeriodo.Substring(0, 4)
        hoja.Cells(8, 9).Value = IdPeriodo.Substring(0, 4)
        hoja.Cells(8, 4).Value = IdPeriodoAnt.Substring(0, 4)
        hoja.Cells(8, 6).Value = IdPeriodoAnt.Substring(0, 4)
        hoja.Cells(8, 8).Value = IdPeriodoAnt.Substring(0, 4)
        hoja.Cells(8, 10).Value = IdPeriodoAnt.Substring(0, 4)
        'GeneraComparativoVentasGrupo(hoja, IdPeriodo)

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

    Private Sub GeneraVentasGrupo(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet
        Dim dt As DataTable

        Dim FechaIni As String, FechaFin As String
        Dim fil_ini As Integer

        Dim array As Object(,)
        Dim Cant As Integer

        fil_ini = 8

        FechaIni = IdPeriodo.Substring(0, 4) & "-01-01"
        FechaFin = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        hoja.Cells(5, 2).Value = "A " & Convert.ToDateTime(FechaFin).ToString("MMMM yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        'hoja.Cells(9, 5).Value = hoja.Cells(9, 4).Value

        hoja.Cells(fil_ini, 5).Value = BuscaVentaNeta("1", FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini, 6).Value = BuscaVentaNeta("2", FechaIni, FechaFin, True, True, False)
        hoja.Cells(fil_ini, 7).Value = BuscaVentaNeta("3", FechaIni, FechaFin, True, True, False)

        hoja.Cells(fil_ini + 11, 4).Value = BuscaVentaNeta("1", FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 12, 4).Value = BuscaVentaNeta("2", FechaIni, FechaFin, True, False, False)
        hoja.Cells(fil_ini + 13, 4).Value = BuscaVentaNeta("3", FechaIni, FechaFin, True, False, False)

        hoja.Cells(fil_ini + 2, 5).Value = -1 * FuncionesEGP.BuscaComisionVolumenEGPAcumulado(IdPeriodo)
        hoja.Cells(fil_ini + 3, 5).Value = -1 * FuncionesVentas.BuscaComisionesInternacionales(IdPeriodo)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select 'Facturación ' + convert(varchar(4), year(FechaDocumento)) as Año, 'US$' as Moneda, -1 * sum(SaldoUSD) as SaldoUSD " & _
                "from LibroInventario " & _
                "where CodCuenta in ('1220001', '122001', '491001') and IdPeriodo = " & IdPeriodo & " " & _
                "group by year(FechaDocumento) order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        fil_ini = 28 + (6 - Cant)
        If Cant < 6 Then hoja.Rows((fil_ini - (6 - Cant)).ToString & ":" & (fil_ini - 1).ToString).EntireRow.Hidden = True

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        sql = "select 'Contratos ' + convert(varchar(4), IdAñoContrato) as Año, 'US$' as Moneda, isnull(sum(SaldoFactLetrasUSD), 0) as SaldoFactLetrasUSD " & _
                "from Consumo " & _
                "where IdPeriodo = " & IdPeriodo & " group by IdAñoContrato having sum(SaldoFactLetrasUSD) > 0"

        cmd = New SqlCommand(sql, cn)
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        fil_ini = 40 + (6 - Cant)
        If Cant < 6 Then hoja.Rows((fil_ini - (6 - Cant)).ToString & ":" & (fil_ini - 1).ToString).EntireRow.Hidden = True

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        cn.Close()
    End Sub

    Private Sub CreaVentaPpto(ByVal IdAño As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_VentaPpto"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdAño", IdAño)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Private Function LlenaVentaPpto(ByVal IdAño As Integer) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPeriodo, Medio, CodRubro, VentaUSD " & _
                "from VentaPpto V, Medio M " & _
                "where V.CodMedio = M.CodMedio and substring(convert(varchar(6), IdPeriodo), 1, 4) = " & IdAño.ToString & " order by 1, 2, 3"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub GeneraPlantillaVentaPpto(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdAño As Integer)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("VentaPpto")
        hoja.Unprotect()

        CreaVentaPpto(IdAño)

        Dim array As Object(,)
        Dim dt As DataTable
        Dim cant As Integer

        dt = LlenaVentaPpto(IdAño)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, False)
        hoja.Range("A2:D" & (cant + 1).ToString, Type.Missing).Value2 = array

        hoja.Protect()

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

    Private Function LlenaVenta2() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Medio, Fecha, CodRubro, Cliente, VentaAcumUSD from Venta2 order by 1, 2, 3, 4"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub GeneraPlantillaVenta2(ByVal RutaPlantilla As String, ByVal RutaArchivo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("Venta2")

        Dim array As Object(,)
        Dim dt As DataTable
        Dim cant As Integer

        dt = LlenaVenta2()
        cant = dt.Rows.Count

        If cant > 0 Then
            array = DataSet2Array(dt, False)
            hoja.Range("A2:E" & (cant + 1).ToString, Type.Missing).Value2 = array
        End If

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

    'Private Sub GeneraComparativoVentasGrupo(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
    '    Dim FechaIni As String, FechaFin As String
    '    Dim fil_ini As Integer

    '    fil_ini = 9

    '    FechaIni = IdPeriodo.Substring(0, 4) & "-01-01"
    '    FechaFin = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

    '    hoja.Cells(5, 2).Value = "A " & Convert.ToDateTime(FechaFin).ToString("MMMM yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

    '    hoja.Cells(fil_ini, 3).Value = BuscaVentaNeta("1", FechaIni, FechaFin, True, True, False)
    '    'hoja.Cells(fil_ini, 5).Value = BuscaVentaNeta("2", FechaIni, FechaFin, True, True, False)
    '    hoja.Cells(fil_ini, 7).Value = BuscaVentaNeta("3", FechaIni, FechaFin, True, True, False)

    '    FechaIni = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1) & "-01-01"
    '    FechaFin = Convert.ToDateTime((Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

    '    hoja.Cells(fil_ini, 4).Value = BuscaVentaNeta("1", FechaIni, FechaFin, True, True, False)
    '    'hoja.Cells(fil_ini, 6).Value = BuscaVentaNeta("2", FechaIni, FechaFin, True, True, False)
    '    hoja.Cells(fil_ini, 8).Value = BuscaVentaNeta("3", FechaIni, FechaFin, True, True, False)
    'End Sub

End Module

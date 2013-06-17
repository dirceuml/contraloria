Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesCtasPorCobrar
    Dim aux_formula As String

    Function LlenaCtasPorCobrar() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaPorCobrar, CtaPorCobrar, Abreviado, CC.CodCuenta, CC.CodCuenta + ' ' + C.Cuenta as Cuenta, CodSeccion, " & _
                "case CodSeccion when 'CLIENT' then 'Clientes' when 'ANTICP' then 'Anticipos' else 'Ctas por Cobrar' end as Seccion, Orden " & _
                "from CtaPorCobrar CC, CuentaEGP C " & _
                "where CC.CodCuenta = C.CodCuenta order by Orden, CC.CodCuenta"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaCuentasEGP_CtaPorCobrar() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCuenta, CodCuenta + ' ' + Cuenta as Cuenta from CuentaEGP " & _
                "where CodCuenta like '1%' or CodCuenta like '422%' order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub CreaCtaPorCobrar(ByVal CtaPorCobrar As String, ByVal Abreviado As String, ByVal CodCuenta As String, ByVal CodSeccion As String, ByVal Orden As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim IdCtaPorCobrar As String = ""

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaPorCobrar from CtaPorCobrar where CtaPorCobrar = '" & IdCtaPorCobrar & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then IdCtaPorCobrar = rdr("IdCtaPorCobrar").ToString()
        rdr.Close()

        If IdCtaPorCobrar = "" Then
            sql = "select isnull(max(IdCtaPorCobrar),0) + 1 as IdCtaPorCobrar from CtaPorCobrar "
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader
            rdr.Read()
            IdCtaPorCobrar = rdr("IdCtaPorCobrar").ToString()
            rdr.Close()
        End If

        sql = "insert CtaPorCobrar(IdCtaPorCobrar, CtaPorCobrar, Abreviado, CodCuenta, CodSeccion, Orden) " & _
                "values (" & IdCtaPorCobrar & ", '" & CtaPorCobrar & "', '" & Abreviado & "', '" & CodCuenta & "', '" & CodSeccion & "', " & Orden & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaPorCobrar(ByVal IdCtaPorCobrar As String, ByVal CodCuenta As String, ByVal CtaPorCobrar As String, ByVal Abreviado As String, ByVal CodCuenta2 As String, ByVal CodSeccion As String, ByVal Orden As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaPorCobrar set CtaPorCobrar = '" & CtaPorCobrar & "', Abreviado = '" & Abreviado & "', " & _
                "CodCuenta = '" & CodCuenta2 & "', CodSeccion = '" & CodSeccion & "' " & _
                "where IdCtaPorCobrar = " & IdCtaPorCobrar & " and CodCuenta = '" & CodCuenta & "' "
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "update CtaPorCobrar set CtaPorCobrar = '" & CtaPorCobrar & "', Abreviado = '" & Abreviado & "', " & _
                "CodSeccion = '" & CodSeccion & "' " & _
                "where IdCtaPorCobrar = " & IdCtaPorCobrar & " "
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaPorCobrar(ByVal IdCtaPorCobrar As String, ByVal CodCuenta As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaPorCobrar " & _
                "where IdCtaPorCobrar = " & IdCtaPorCobrar & " and CodCuenta = '" & CodCuenta & "' "
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function BuscaCtaPorCobrar(ByVal IdCtaPorCobrar As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim CtaPorCobrar As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select top 1 CtaPorCobrar from CtaPorCobrar where IdCtaPorCobrar = " & IdCtaPorCobrar
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        CtaPorCobrar = rdr("CtaPorCobrar").ToString()
        rdr.Close()

        cn.Close()

        Return CtaPorCobrar
    End Function

    Function BuscaAbreviado(ByVal IdCtaPorCobrar As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Abreviado As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select top 1 Abreviado from CtaPorCobrar where IdCtaPorCobrar = " & IdCtaPorCobrar
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Abreviado = rdr("Abreviado").ToString()
        rdr.Close()

        cn.Close()

        Return Abreviado
    End Function

    Function BuscaCtasPorCobrar() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct IdCtaPorCobrar, CtaPorCobrar, Orden from CtaPorCobrar order by Orden desc"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaFacturasPorCobrar(ByVal CodSeccion As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        'If CodSeccion = "FACTUR" Then
        '    sql = "select distinct year(FechaDocumento) * 100 + month(FechaDocumento) as IdPeriodo from V_FacturaPorCobrar " & _
        '            "where year(FechaDocumento) * 100 + month(FechaDocumento) <= " & IdPeriodo & " and CodSeccion = '" & CodSeccion & "' "
        'ElseIf CodSeccion = "LETRAS" Then
        '    sql = "select distinct year(FechaDocumento) * 100 + month(FechaDocumento) as IdPeriodo from V_FacturaPorCobrar " & _
        '            "where CodSeccion = '" & CodSeccion & "' "
        'Else
        '    sql = "select distinct year(FechaDocumento) as IdPeriodo from V_FacturaPorCobrar " & _
        '            "where year(FechaDocumento) * month(FechaDocumento) <= " & IdPeriodo & " and CodSeccion = '" & CodSeccion & "' "
        'End If

        If CodSeccion = "FACTUR" Or CodSeccion = "LETRAS" Then
            sql = "select distinct year(FechaDocumento) * 100 + month(FechaDocumento) as IdPeriodo from FacturaPorCobrar " & _
                    "where CodSeccion = '" & CodSeccion & "' "
        Else
            sql = "select distinct year(FechaDocumento) as IdPeriodo from FacturaPorCobrar " & _
                    "where CodSeccion = '" & CodSeccion & "' "
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub CreaFacturaPorCobrar(ByVal Fecha As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_FacturaPorCobrar"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("Fecha", Fecha)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraReportesCtasPorCobrar(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String, ByVal flagContraloria As Boolean)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet
        Dim dt As DataTable
        Dim UltDiaMes As Date
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("RESUMEN")
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(5, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        reporte.Worksheets("CARATULA").Cells(45, 4).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(8, 3).Value = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & "-12-31"
        hoja.Cells(8, 5).Value = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        GeneraResumen(hoja, IdPeriodo, "CLIENT", 10, flagContraloria)
        GeneraResumen(hoja, IdPeriodo, "CTASXC", 24, flagContraloria)
        GeneraResumen(hoja, IdPeriodo, "ANTICP", 38, flagContraloria)
        dt = BuscaCtasPorCobrar()
        For Each row In dt.Rows
            GeneraDetalle(reporte, IdPeriodo, row("IdCtaPorCobrar").ToString, 50, 1, flagContraloria)
        Next
        reporte.Worksheets("Plantilla").Delete()

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

    Sub GeneraResumen(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal flagContraloria As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim tabla As String

        Dim IdAño, IdAñoAnt As String

        Dim array As Object(,)
        Dim cant As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not flagContraloria Then tabla = "V_CtaPorCobrar" Else tabla = "V_CtaPorCobrar_Contraloria"

        sql = "select upper(CtaPorCobrar) as CtaPorCobrar, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto from " & _
                "(select IdCtaPorCobrar, sum(MontoUSD) as MontoAnt " & _
                "from " & tabla & " where IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 " & _
                "group by IdCtaPorCobrar) C1 full join " & _
                "(select IdCtaPorCobrar, sum(MontoUSD) as Monto " & _
                "from " & tabla & " where IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by IdCtaPorCobrar) C2 on C1.IdCtaPorCobrar = C2.IdCtaPorCobrar full join " & _
                "(select distinct IdCtaPorCobrar, CtaPorCobrar, CodSeccion, Orden from CtaPorCobrar) C on isnull(C1.IdCtaPorCobrar, C2.IdCtaPorCobrar) = C.IdCtaPorCobrar " & _
                "where CodSeccion = '" & CodSeccion & "' order by Orden"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 0, 1, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 2, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + cant <= fil_ini + 12 - 1 And CodSeccion <> "ANTICP" Then
            hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + 12 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraDetalle(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal IdCtaPorCobrar As String, ByVal CantReg As Integer, ByVal fil_ini As Integer, ByVal flagContraloria As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim tabla As String

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim IdAño, IdAñoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, i, i2 As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        reporte.Worksheets("Plantilla").Copy(, reporte.Worksheets("Plantilla"))
        hoja = reporte.Worksheets("Plantilla (2)")
        hoja.Name = BuscaAbreviado(IdCtaPorCobrar).ToUpper

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not flagContraloria Then tabla = "V_CtaPorCobrar" Else tabla = "V_CtaPorCobrar_Contraloria"

        sql = "select Persona, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto from " & _
                "(select CodPersona, sum(MontoUSD) as MontoAnt " & _
                "from " & tabla & " where IdCtaPorCobrar = " & IdCtaPorCobrar & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 " & _
                "group by CodPersona) PA1 full join " & _
                "(select CodPersona, sum(MontoUSD) as Monto " & _
                "from " & tabla &" where IdCtaPorCobrar = " & IdCtaPorCobrar & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodPersona) PA2 on PA1.CodPersona = PA2.CodPersona join " & _
                "Persona P on isnull(PA1.CodPersona, PA2.CodPersona) = P.CodPersona " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by Persona"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        i = 0
        If cant = 0 Then
            hoja.Rows("11:59").EntireRow.Hidden = True
            hoja.Cells(fil_ini + 3, 2).Value = BuscaCtaPorCobrar(IdCtaPorCobrar).ToUpper
            UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        End If

        Do While i < cant
            rango = reporte.Worksheets("Plantilla").Rows("1:61")
            rango.Copy(hoja.Rows(fil_ini))

            hoja.Cells(fil_ini + 3, 2).Value = BuscaCtaPorCobrar(IdCtaPorCobrar).ToUpper
            UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            If i > 0 Then
                hoja.Rows((fil_ini + 2).ToString & ":" & (fil_ini + 4).ToString).EntireRow.Hidden = True
                hoja.Cells(fil_ini + 8, 4).FormulaR1C1 = "=R[-10]C[0]"
                hoja.Cells(fil_ini + 8, 5).FormulaR1C1 = "=R[-10]C[0]"
                hoja.Cells(fil_ini + 8, 6).FormulaR1C1 = "=R[-10]C[0]"
            Else
                hoja.Rows((fil_ini + 8).ToString & ":" & (fil_ini + 8).ToString).EntireRow.Hidden = True
            End If
            hoja.Cells(fil_ini + 7, 4).Value = IdAñoAnt & "-12-31"
            hoja.Cells(fil_ini + 7, 6).Value = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

            fil_ini = fil_ini + 9
            hoja.Cells(fil_ini - 1, 2).Value = i
            If i + CantReg < cant Then i2 = i + CantReg - 1 Else i2 = cant - 1
            array = DataSet2Array(dt, i, i2, 2, 0, 1, -1, -1, -1)
            hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dt, i, i2, 1, 2, -1, -1, -1, -1)
            hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array

            If i2 = cant - 1 Then
                If cant Mod CantReg > 0 Then hoja.Rows((fil_ini + i2 - i + 1).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True
                hoja.Cells(fil_ini + CantReg, 3).Value = "TOTAL " & BuscaCtaPorCobrar(IdCtaPorCobrar).ToUpper & " US$"
                hoja.PageSetup.PrintArea = "$B$1:$F$" + CStr(fil_ini + CantReg)
            End If
            fil_ini = fil_ini + i2 - i + 3
            i = i + CantReg
            If i2 <> cant - 1 Then hoja.HPageBreaks.Add(hoja.Rows(fil_ini - 1))
        Loop

        cn.Close()
    End Sub

    Sub GeneraReportesFacturasPorCobrar(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal Fecha As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        CreaFacturaPorCobrar(Fecha)

        GeneraFacturas(reporte, Fecha)

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

    Sub GeneraFacturas(ByVal reporte As Excel.Workbook, ByVal Fecha As String)
        Dim hoja As Excel.Worksheet, rango As Excel.Range
        Dim dt As DataTable
        Dim IdPeriodoAux As String
        Dim fil_ini As Integer

        aux_formula = "="
        hoja = reporte.Worksheets("FACTURAS")

        'Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(3, 2).Value = "AL " & Convert.ToDateTime(Fecha).ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        dt = BuscaFacturasPorCobrar("FACTUR")
        fil_ini = 10
        For Each row In dt.Rows
            IdPeriodoAux = row("IdPeriodo")
            If IdPeriodoAux = Convert.ToDateTime(Fecha).ToString("yyyyMM") Then
                GeneraDetalleFacturas(hoja, IdPeriodoAux, Fecha, "FACTUR", fil_ini)
            Else
                GeneraDetalleFacturas(hoja, IdPeriodoAux, "", "FACTUR", fil_ini)
            End If
        Next
        hoja.Rows("6:9").EntireRow.Hidden = True
        rango = hoja.Rows("8:8")
        rango.Copy(hoja.Rows(fil_ini))
        hoja.Cells(fil_ini, 7).FormulaR1C1 = aux_formula
        hoja.Cells(fil_ini, 8).FormulaR1C1 = aux_formula
        hoja.Rows(fil_ini).EntireRow.Hidden = False
        hoja.PageSetup.PrintArea = "$B$1:$H$" + CStr(fil_ini)

        aux_formula = "="
        hoja = reporte.Worksheets("LETRAS")

        hoja.Cells(3, 2).Value = "AL " & Convert.ToDateTime(Fecha).ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper


        dt = BuscaFacturasPorCobrar("LETRAS")
        fil_ini = 10
        For Each row In dt.Rows
            IdPeriodoAux = row("IdPeriodo")
            If IdPeriodoAux = Convert.ToDateTime(Fecha).ToString("yyyyMM") Then
                GeneraDetalleFacturas(hoja, IdPeriodoAux, Fecha, "LETRAS", fil_ini)
            Else
                GeneraDetalleFacturas(hoja, IdPeriodoAux, "", "LETRAS", fil_ini)
            End If
        Next
        hoja.Rows("6:9").EntireRow.Hidden = True
        rango = hoja.Rows("8:8")
        rango.Copy(hoja.Rows(fil_ini))
        hoja.Cells(fil_ini, 7).FormulaR1C1 = aux_formula
        hoja.Cells(fil_ini, 8).FormulaR1C1 = aux_formula
        hoja.Rows(fil_ini).EntireRow.Hidden = False
        hoja.PageSetup.PrintArea = "$B$1:$H$" + CStr(fil_ini)

        aux_formula = "="
        hoja = reporte.Worksheets("FACTURAS VARIAS")

        hoja.Cells(3, 2).Value = "AL " & Convert.ToDateTime(Fecha).ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        dt = BuscaFacturasPorCobrar("FACVAR")
        fil_ini = 10
        For Each row In dt.Rows
            IdPeriodoAux = row("IdPeriodo")
            If IdPeriodoAux = Convert.ToDateTime(Fecha).ToString("yyyyMM") Then
                GeneraDetalleFacturas(hoja, IdPeriodoAux, Fecha, "FACVAR", fil_ini)
            Else
                GeneraDetalleFacturas(hoja, IdPeriodoAux, "", "FACVAR", fil_ini)
            End If
        Next
        hoja.Rows("6:9").EntireRow.Hidden = True
        rango = hoja.Rows("8:8")
        rango.Copy(hoja.Rows(fil_ini))
        hoja.Cells(fil_ini, 7).FormulaR1C1 = aux_formula
        hoja.Cells(fil_ini, 8).FormulaR1C1 = aux_formula
        hoja.Rows(fil_ini).EntireRow.Hidden = False
        hoja.PageSetup.PrintArea = "$B$1:$H$" + CStr(fil_ini)
    End Sub

    Sub GeneraDetalleFacturas(ByVal hoja As Excel.Worksheet, ByVal IdPeriodoAux As String, ByVal Fecha As String, ByVal CodSeccion As String, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim auxsql As String
        Dim rango As Excel.Range

        Dim array As Object(,)

        Dim cant As Integer

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        'If IdPeriodoAux <> IdPeriodo.Substring(0, 4) Then IdPeriodo = IdPeriodoAux & "12"

        If Fecha <> "" Then auxsql = "and FechaDocumento <= '" & Fecha & "' " Else auxsql = ""

        If CodSeccion = "FACTUR" Then
            sql = "select FechaDocumento, Serie, Documento, Persona, Contrato, SaldoDolares, SaldoSoles " & _
                    "from FacturaPorCobrar " & _
                    "where year(FechaDocumento) * 100 + month(FechaDocumento) = " & IdPeriodoAux & " " & auxsql & _
                    "and CodSeccion = '" & CodSeccion & "' order by FechaDocumento, Serie, Documento"
        ElseIf CodSeccion = "LETRAS" Then
            sql = "select Serie, Documento, FechaDocumento, Persona, Contrato, SaldoDolares, SaldoSoles " & _
                    "from FacturaPorCobrar " & _
                    "where year(FechaDocumento) * 100 + month(FechaDocumento) = " & IdPeriodoAux & " " & auxsql & _
                    "and CodSeccion = '" & CodSeccion & "' order by FechaDocumento, Serie, Documento"
        Else
            sql = "select FechaDocumento, Serie, Documento, Persona, null as Contrato, SaldoDolares, SaldoSoles " & _
                    "from FacturaPorCobrar " & _
                    "where year(FechaDocumento) = " & IdPeriodoAux & " " & auxsql & _
                    "and CodSeccion = '" & CodSeccion & "' order by FechaDocumento, Serie, Documento"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        rango = hoja.Rows("6:6")
        rango.Copy(hoja.Rows((fil_ini).ToString & ":" & (fil_ini + cant - 1)))
        'hoja.Rows((fil_ini).ToString & ":" & (fil_ini + cant - 1)).EntireRow.Insert()
        array = DataSet2Array(dt, False)
        hoja.Range("B" & (fil_ini).ToString & ":H" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        rango = hoja.Rows("7:7")
        rango.Copy(hoja.Rows(fil_ini + cant))

        If CodSeccion = "FACTUR" Or CodSeccion = "LETRAS" Then
            Dim aux As Date
            aux = Convert.ToDateTime(IdPeriodoAux.Substring(0, 4) & "-" & IdPeriodoAux.Substring(4, 2) & "-01")
            hoja.Cells(fil_ini + cant, 5).Value = "TOTAL " & aux.ToString("MMMM yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        Else
            hoja.Cells(fil_ini + cant, 5).Value = "TOTAL AÑO " & IdPeriodoAux
        End If
        hoja.Cells(fil_ini + cant, 7).FormulaR1C1 = "=SUM(R[-" & cant.ToString & "]C[0]:R[-1]C[0])"
        hoja.Cells(fil_ini + cant, 8).FormulaR1C1 = "=SUM(R[-" & cant.ToString & "]C[0]:R[-1]C[0])"
        aux_formula = aux_formula & "+ R" & (fil_ini + cant).ToString & "C[0]"
        hoja.Rows(fil_ini + cant + 1).EntireRow.Hidden = True
        fil_ini = fil_ini + cant + 2

        cn.Close()
    End Sub

End Module

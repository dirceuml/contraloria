Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesEvalEGP

    Private Sub CreaEGPPpto(ByVal IdAño As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_EGPPpto"
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

    Private Function LlenaEGPPptoOtrosRubros(ByVal IdAño As Integer) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select E.IdAño, 0 as CodCentroCosto, 'Sin Centro Costo' as CentroCosto, C.CodSeccion, E.IdCtaEGP, C.CtaEGP, " & _
                "Ene, Feb, Mar, Abr, May, Jun, Jul, Ago, [Set], Oct, Nov, Dic " & _
                "from EGPPpto E, CtaEGP C " & _
                "where E.IdCtaEGP = C.IdCtaEGP and C.CodSeccion <> 'COSTOS' and IdAño = " & IdAño.ToString & " order by E.IdAño, C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Private Function LlenaEGPPptoCostos(ByVal IdAño As Integer) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select E.IdAño, E.CodCentroCosto, CC.CentroCostoEGP, C.CodSeccion, E.IdCtaEGP, C.CtaEGP, " & _
                "Ene, Feb, Mar, Abr, May, Jun, Jul, Ago, [Set], Oct, Nov, Dic " & _
                "from EGPPpto E, CtaEGP C, CentroCosto CC " & _
                "where E.IdCtaEGP = C.IdCtaEGP and E.CodCentroCosto = CC.CodCentroCosto " & _
                "and C.CodSeccion = 'COSTOS' and E.IdAño = " & IdAño.ToString & " order by E.IdAño, CC.CodCentroCosto, C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub GeneraPlantillaEGPPpto(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdAño As Integer)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet

        Dim array As Object(,)
        Dim dt As DataTable
        Dim cant As Integer

        'Try

        CreaEGPPpto(IdAño)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("Otros Rubros")
        hoja.Unprotect()

        dt = LlenaEGPPptoOtrosRubros(IdAño)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, False)
        hoja.Range("A2:R" & (cant + 1).ToString, Type.Missing).Value2 = array

        hoja.Protect(, , , , , , , , , , , , , True, True)

        hoja = reporte.Worksheets("Costos")
        hoja.Unprotect()

        dt = LlenaEGPPptoCostos(IdAño)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, False)
        hoja.Range("A2:R" & (cant + 1).ToString, Type.Missing).Value2 = array

        hoja.Protect(, , , , , , , , , , , , , True, True)

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

    Sub CreaDetalleEGPPpto(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        Dim IdAño As String

        IdAño = IdPeriodo.Substring(0, 4)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "sp_Crea_DetalleEGPPpto"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("IdAño", Convert.ToInt32(IdAño))
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub GeneraReportesEGPPpto(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdAño As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        Dim IdPeriodo As String
        Dim IdAñoMes As Integer

        'Try
        IdPeriodo = IdAño & "12"
        FuncionesEGP.CreaDetalleEGP(IdPeriodo)
        CreaDetalleEGPPpto(IdPeriodo)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("CARATULA")
        hoja.Cells(8, 2).Value = "AÑO " & IdAño.ToString

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PPTO x MESES")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, False, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, False, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PPTO")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PPTO x MESES v2")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, True, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, True, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PPTO")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("COSTOS PPTO x RUBROS")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosCostosMes(hoja, IdPeriodo, IdAñoMes, 8, 67, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 70, 79, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "26", IdAñoMes, 82, 82, 3, True, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "DEP", IdAñoMes, 85, 99, 3, True, "PPTO")
        Next

        GeneraCostosPpto(reporte, IdAño, "I", False)
        GeneraCostosPpto(reporte, IdAño, "I", True)
        GeneraCostosPpto(reporte, IdAño, "D", False)

        reporte.Worksheets("Plantilla COSTOS").Delete()

        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(reporte)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        hoja = Nothing
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

    Sub GeneraCostosPpto(ByVal reporte As Excel.Workbook, ByVal IdAño As String, ByVal CodTipo As String, ByVal FlagAgrupado As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String

        Dim hoja As Excel.Worksheet
        Dim IdAñoMes As Integer
        Dim fil_ini, i As Integer

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not FlagAgrupado Then
            sql = "select distinct E.CodCentroCosto, GrupoEGP, CentroCostoEGP " & _
                    "from DetalleEGP E, CentroCosto C where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' " & _
                    "and CodTipo = '" & CodTipo & "' and IdPeriodo between " & IdAño & "01 and " & IdAño & "12 " & _
                    "order by 2, 3"
        Else
            sql = "select distinct CodGrupoEGP as CodCentroCosto, GrupoEGP " & _
                    "from DetalleEGP E, CentroCosto C where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and CodTipo = 'I' and CodGrupoEGP in (5, 6) " & _
                    "and IdPeriodo between " & IdAño & "01 and " & IdAño & "12 " & _
                    "order by 2"
        End If

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader

        reporte.Worksheets("Plantilla COSTOS").Copy(, reporte.Worksheets("Plantilla COSTOS"))
        hoja = reporte.Worksheets("Plantilla COSTOS (2)")
        hoja.Name = "-"

        fil_ini = 2
        i = 0
        While rdr.Read
            If Not FlagAgrupado Then
                Dim CentroCostoEGP = rdr("CentroCostoEGP").ToString.Replace("Costos ", "").ToUpper
                If CentroCostoEGP.Length > 15 Then CentroCostoEGP = CentroCostoEGP.Substring(0, 15)
                If hoja.Name = "-" Then
                    If CodTipo = "D" Then hoja.Name = CentroCostoEGP & "-" Else hoja.Name = CentroCostoEGP
                Else
                    hoja.Name = hoja.Name & CentroCostoEGP
                End If
            Else
                Dim GrupoEGP = rdr("GrupoEGP").ToString.Replace("Costos ", "").ToUpper
                hoja.Name = GrupoEGP & " AGRUPADO"
            End If

            IdAñoMes = IdAño * 100
            For j = 1 To 12
                IdAñoMes = IdAñoMes + 1
                GeneraCostosPorCentroMesPpto(hoja, (IdAño * 100 + 12).ToString, rdr("CodCentroCosto"), IdAñoMes.ToString, fil_ini, False)
            Next
            fil_ini = fil_ini + 47

            i = i + 1
            If i Mod 2 = 0 Or CodTipo = "I" Or FlagAgrupado Then
                If CodTipo = "I" Or FlagAgrupado Then hoja.Rows("48:94").EntireRow.Hidden = True
                reporte.Worksheets("Plantilla COSTOS").Copy(, hoja)
                hoja = reporte.Worksheets("Plantilla COSTOS (2)")
                hoja.Name = "-"
                fil_ini = 2
            End If
        End While
        If hoja.Name = "-" Then hoja.Delete()

        rdr.Close()
        cn.Close()
    End Sub

    Sub GeneraCostosPorCentroMesPpto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodCentroCosto As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal FlagAgrupado As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim array As Object(,)
        Dim col_ini, fil_fin As Integer
        Dim IdAño As String
        Dim IdMesAux As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        hoja.Cells(fil_ini, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        hoja.Cells(fil_ini + 1, 3).Value = BuscaCentroCosto(CodCentroCosto).ToUpper

        col_ini = 3
        fil_ini = fil_ini + 5
        fil_fin = fil_ini + 40 - 1

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not FlagAgrupado Then
            sql = "select E.CtaEGP, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.PptoUSD, 0) as PptoAcum " &
                    "from (select IdCtaEGP, CtaEGP from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                    "(select CodCentroCosto, IdCtaEGP, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
                    "where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodoAux & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C " & _
                    "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                    "(select CodCentroCosto, IdCtaEGP, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
                    "where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C1 " & _
                    "on E.IdCtaEGP = C1.IdCtaEGP " & _
                    "where not C1.PptoUSD is null " & _
                    "order by 1"
        Else
            sql = "select E.CtaEGP, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.PptoUSD, 0) as PptoAcum " &
                    "from (select IdCtaEGP, CtaEGP from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                    "(select IdCtaEGP, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                    "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodoAux & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP ) C " & _
                    "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                    "(select IdCtaEGP, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                    "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP) C1 " & _
                    "on E.IdCtaEGP = C1.IdCtaEGP " & _
                    "where not C1.PptoUSD is null " & _
                    "order by 1"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        If IdMesAux = 1 Then
            array = DataSet2Array(dtaset.Tables(0), 1, 0, -1, -1, -1, -1)
            hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini)).Value2 = array
        End If

        array = DataSet2Array(dtaset.Tables(0), 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMesAux), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini + IdMesAux)).Value2 = array

        If IdMesAux = 12 And fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraReportesEGPProy(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        Dim IdAño, IdMes, IdAñoMes As Integer

        'Try

        CreaDetalleEGPPpto(IdPeriodo)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("CARATULA")
        hoja.Cells(8, 2).Value = "A " & BuscaPeriodo(IdPeriodo).ToUpper

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        hoja = reporte.Worksheets("EGP EJEC vs PPTO")
        hoja.Cells(5, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(7, 3).Value = BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(7, 7).Value = "ACUMULADO " & IdAño

        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGPUB", 10, 11)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "COMVOL", 13, 13)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGVEN", 15, 15)
        GeneraCentrosCostoEjecPpto(hoja, IdPeriodo, "D", 19, 68, False)
        GeneraCentrosCostoEjecPpto(hoja, IdPeriodo, "I", 71, 90, False)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "OTRCOS", 93, 102)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "OTRRUB", 108, 122)

        hoja = reporte.Worksheets("EGP EJEC vs PPTO v2")
        hoja.Cells(5, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(7, 3).Value = BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(7, 7).Value = "ACUMULADO " & IdAño

        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGPUB", 10, 11)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "COMVOL", 13, 13)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGVEN", 15, 15)
        GeneraCentrosCostoEjecPpto(hoja, IdPeriodo, "D", 19, 68, True)
        GeneraCentrosCostoEjecPpto(hoja, IdPeriodo, "I", 71, 90, True)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "OTRCOS", 93, 102)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "OTRRUB", 108, 122)

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PPTO x MESES")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, False, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, False, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PPTO")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PPTO x MESES v2")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, True, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, True, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PPTO")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PROY x MESES")
        hoja.Cells(2, 3).Value = "PROYECTADO " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PROYECTADO " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            If IdAñoMes <= IdPeriodo Then
                hoja.Cells(6, 3 + i).Value = "EJEC"
                hoja.Cells(107, 3 + i).Value = "EJEC"
            End If
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, False, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, False, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PROY")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("EGP PROY x MESES v2")
        hoja.Cells(2, 3).Value = "PROYECTADO " & IdAño.ToString
        hoja.Cells(103, 3).Value = "PROYECTADO " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            If IdAñoMes <= IdPeriodo Then
                hoja.Cells(6, 3 + i).Value = "EJEC"
                hoja.Cells(107, 3 + i).Value = "EJEC"
            End If
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 8, 9, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 11, 11, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 13, 13, 3, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 16, 64, 3, True, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 67, 86, 3, True, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 89, 98, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 109, 123, 3, "PROY")
        Next

        hoja = reporte.Worksheets("INGRESOS")
        hoja.Cells(11, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(13, 3).Value = BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(13, 7).Value = "ACUMULADO " & IdAño
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGPUB", 16, 17)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "COMVOL", 19, 19)
        GeneraRubrosEjecPpto(hoja, IdPeriodo, "INGVEN", 21, 21)

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("COSTOS PPTO x RUBROS")
        hoja.Cells(2, 3).Value = "PRESUPUESTO ANUAL " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosCostosMes(hoja, IdPeriodo, IdAñoMes, 8, 67, 3, "PPTO")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 70, 79, 3, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "26", IdAñoMes, 82, 82, 3, True, "PPTO")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "DEP", IdAñoMes, 85, 99, 3, True, "PPTO")
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("COSTOS PROY x RUBROS")
        hoja.Cells(2, 3).Value = "PROYECTADO " & IdAño.ToString
        For i = 1 To 12
            IdAñoMes = IdAñoMes + 1
            If IdAñoMes <= IdPeriodo Then hoja.Cells(6, 3 + i).Value = "EJEC"
            GeneraRubrosCostosMes(hoja, IdPeriodo, IdAñoMes, 8, 67, 3, "PROY")
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 70, 79, 3, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "26", IdAñoMes, 82, 82, 3, True, "PROY")
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "DEP", IdAñoMes, 85, 99, 3, True, "PROY")
        Next

        GeneraCostos(reporte, IdPeriodo, "I", True)
        GeneraCostosAgrupados(reporte, IdPeriodo, True)
        GeneraCostos(reporte, IdPeriodo, "D", True)

        reporte.Worksheets("Plantilla Costos").Delete()
        reporte.Worksheets("COSTOS").Delete()

        reporte.SaveAs(RutaArchivo)
        reporte.Close()
        excel.Quit()
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(reporte)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        hoja = Nothing
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

    Sub GeneraRubrosEjecPpto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMes, IdAño As String
        Dim array As Object(,)

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.MontoUSD, 0) as MontoAcum, isnull(C1.PptoUSD, 0) as PptoAcum, E.Orden " & _
            "from (select IdCtaEGP, CtaEGP, Orden from CtaEGP where CodSeccion ='" & CodSeccion & "') E left join " & _
            "(select IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
            "where CodSeccion = '" & CodSeccion & "' and IdPeriodo = " & IdPeriodo & " group by IdCtaEGP) C " & _
            "on E.IdCtaEGP = C.IdCtaEGP left join " & _
            "(select IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
            "where CodSeccion = '" & CodSeccion & "' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaEGP) C1 " & _
            "on E.IdCtaEGP = C1.IdCtaEGP " & _
            "where (not C1.MontoUSD is null or not C1.PptoUSD is null) " & _
            "order by E.Orden"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 3, 0, 1, 2, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 2, 3, 4, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_fin Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraCentrosCostoEjecPpto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal FlagVersion2 As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMes, IdAño As String

        Dim array As Object(,)

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagVersion2 Then
            sql = "select CC.CentroCostoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.MontoUSD, 0) as CostoAcum, isnull(C1.PptoUSD, 0) as PptoAcum " & _
                "from (select * from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
                "where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " group by CodCentroCosto) C " & _
                "on CC.CodCentroCosto = C.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP  " & _
                "where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) C1 " & _
                "on CC.CodCentroCosto = C1.CodCentroCosto " & _
                "where (not C1.MontoUSD is null or not C1.PptoUSD is null) " & _
                "order by 1"
        Else
            sql = "select CC.GrupoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.MontoUSD, 0) as CostoAcum, isnull(C1.PptoUSD, 0) as PptoAcum " & _
                "from (select distinct CodGrupoEGP, GrupoEGP from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " group by CodGrupoEGP) C " & _
                "on CC.CodGrupoEGP = C.CodGrupoEGP left join " & _
                "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodGrupoEGP) C1 " & _
                "on CC.CodGrupoEGP = C1.CodGrupoEGP " & _
                "where (not C1.MontoUSD is null or not C1.PptoUSD is null) " & _
                "order by 1"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 3, 0, 1, 2, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 2, 3, 4, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array

        hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

    End Sub

    Sub GeneraCostosPorCentroEjecPpto(ByVal reporte As Excel.Workbook, ByVal hoja As Excel.Worksheet, ByVal CodCentroCosto As String, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal FlagAgrupado As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim rango As Excel.Range

        Dim array As Object(,)

        Dim IdAño As String

        IdAño = IdPeriodo.Substring(0, 4)

        rango = reporte.Worksheets("Plantilla Costos").Rows("4:50")
        rango.Copy(hoja.Rows(fil_ini))

        hoja.Cells(fil_ini, 2).Value = BuscaCentroCosto(CodCentroCosto).ToUpper & " - EJECUTADO vs PRESUPUESTO"
        hoja.Cells(fil_ini + 1, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(fil_ini + 3, 3).Value = BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(fil_ini + 3, 7).Value = "ACUMULADO " & IdAño

        fil_ini = fil_ini + 5

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not FlagAgrupado Then
            sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.MontoUSD, 0) as CostoAcum, isnull(C1.PptoUSD, 0) as PptoAcum " &
                    "from (select * from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                    "(select CodCentroCosto, IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
                    "where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C " & _
                    "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                    "(select CodCentroCosto, IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP " & _
                    "where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C1 " & _
                    "on E.IdCtaEGP = C1.IdCtaEGP " & _
                    "where (not C1.MontoUSD is null or not C1.PptoUSD is null) " & _
                    "order by 1"
        Else
            sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C.PptoUSD, 0) as PptoMes, isnull(C1.MontoUSD, 0) as CostoAcum, isnull(C1.PptoUSD, 0) as PptoAcum " &
                    "from (select * from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                    "(select IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                    "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP ) C " & _
                    "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                    "(select IdCtaEGP, sum(MontoUSD) as MontoUSD, sum(PptoUSD) as PptoUSD from DetalleEGP D, CentroCosto C " & _
                    "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP) C1 " & _
                    "on E.IdCtaEGP = C1.IdCtaEGP " & _
                    "where (not C1.MontoUSD is null or not C1.PptoUSD is null) " & _
                    "order by 1"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 3, 0, 1, 2, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 2, 3, 4, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_ini + 40 - 1 Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & (fil_ini + 40 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

End Module

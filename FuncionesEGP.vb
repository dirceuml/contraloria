Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesEGP

    Function LlenaEGPExp() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Orden, CtaEGP, CodSeccion, CodCuenta, Cuenta, " & _
                "case D.Signo when 1 then '+' when -1 then '-' else '' end as Signo " & _
                "from CtaEGP C left join CtaEGPDet D on C.IdCtaEGP = D.IdCtaEGP left join CuentaEGP CC on D.CodCtaOrigen = CC.CodCuenta " & _
                "order by Orden, CtaEGP, CodCuenta"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaCentroCosto(ByVal CodCentroCosto As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim CentroCosto As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Convert.ToInt32(CodCentroCosto) < 1000 Then
            sql = "select CentroCostoEGP from CentroCosto where CodCentroCosto = " & CodCentroCosto
        Else
            Dim CodGrupoEGP As String = (Convert.ToInt32(CodCentroCosto) - 1000).ToString
            sql = "select top 1 GrupoEGP as CentroCostoEGP from CentroCosto where CodGrupoEGP = " & CodGrupoEGP
        End If

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        CentroCosto = rdr("CentroCostoEGP").ToString()
        rdr.Close()
        cn.Close()

        Return CentroCosto
    End Function

    Function BuscaRubroCosto(ByVal IdCtaEGP As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim RubroCosto As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()
        sql = "select CtaEGP from CtaEGP where IdCtaEGP = " & IdCtaEGP
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        RubroCosto = rdr("CtaEGP").ToString()
        rdr.Close()
        cn.Close()

        Return RubroCosto
    End Function

    Function BuscaComisionVolumenEGPAcumulado(ByVal IdPeriodo As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Fecha As String
        Dim ComisionVolumen As Double

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        Fecha = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        sql = "select sum(DebeUSD) - sum(HaberUSD) as ComisionVolumen " & _
                "from V_Asiento where CodCuenta = '953001102'" & _
                "and Fecha between '" & IdPeriodo.Substring(0, 4) & "-01-01' and '" & Fecha & "' "

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        ComisionVolumen = Convert.ToDouble(rdr("ComisionVolumen"))
        rdr.Close()
        cn.Close()

        Return ComisionVolumen
    End Function

    Function CargaCentrosCosto() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCentroCosto, CentroCostoEGP  from CentroCosto where CodTipo in ('D', 'I') and CentroCostoEGP <> ''" & _
                "union " & _
                "select 1000 + CodGrupoEGP, GrupoEGP + ' Agrupado' from CentroCosto where CentroCostoEGP <> '' " & _
                "group by CodGrupoEGP, GrupoEGP having count(*) > 1 " & _
                "order by 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function CargaRubrosCosto() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()
        sql = "select * from CtaEGP where CodSeccion = 'COSTOS' order by CtaEGP"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub Titulos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal Titulo As String, ByVal fil_ini As Integer)
        Dim IdAño, IdAñoAnt As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        If (Titulo <> "") Then hoja.Cells(fil_ini, 2).Value = Titulo
        hoja.Cells(fil_ini + 1, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(fil_ini + 3, 3).Value = BuscaPeriodo(IdPeriodo).ToUpper
        hoja.Cells(fil_ini + 3, 5).Value = "ACUMULADO " & IdAño
        hoja.Cells(fil_ini + 3, 6).Value = "PROMEDIO " & IdAño
        hoja.Cells(fil_ini + 3, 7).Value = "ANUAL " & IdAñoAnt
        hoja.Cells(fil_ini + 3, 8).Value = "PROMEDIO " & IdAñoAnt
    End Sub

    Sub TitulosComp(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim IdAño, IdAñoAnt, IdMes As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        IdMes = IdPeriodo.Substring(4, 2)
        hoja.Cells(fil_ini + 1, 2).Value = "ANALISIS ACUMULADO " & BuscaMes(IdMes) & " " & IdAño & " VS. " & BuscaMes(IdMes) & " " & IdAñoAnt
        hoja.Cells(fil_ini + 3, 3).Value = "A " & BuscaMes(IdMes) & " " & IdAño
        hoja.Cells(fil_ini + 3, 4).Value = "A " & BuscaMes(IdMes) & " " & IdAñoAnt
        hoja.Cells(fil_ini + 3, 5).Value = "PROMEDIO " & IdAño
        hoja.Cells(fil_ini + 3, 6).Value = "PROMEDIO " & IdAñoAnt
    End Sub

    Sub ProcesoCargaEGP()
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_ProcesoCargaEGP"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub CreaDetalleEGP(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        Dim IdAño, IdMes As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "sp_Crea_DetalleEGP"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IdAño", IdAño)
        cmd.Parameters.AddWithValue("@IdMes", IdMes)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub GeneraReportesEGP(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        Dim IdAño, IdMes, IdAñoMes As Integer

        'Try

        CreaDetalleEGP(IdPeriodo)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("CARATULA")
        hoja.Cells(7, 2).Value = "A " & BuscaPeriodo(IdPeriodo).ToUpper

        hoja = reporte.Worksheets("EST.GAN.Y PER.")
        GeneraEGP(hoja, IdPeriodo, 4, False, False)

        hoja = reporte.Worksheets("EST.GAN.Y PER. V2")
        GeneraEGP(hoja, IdPeriodo, 4, False, True)

        hoja = reporte.Worksheets("COMPAR.MES")
        GeneraEGP(hoja, IdPeriodo, 4, True, False)

        hoja = reporte.Worksheets("COMPAR.MES V2")
        GeneraEGP(hoja, IdPeriodo, 4, True, True)

        hoja = reporte.Worksheets("INGRESOS")
        Titulos(hoja, IdPeriodo, "INGRESOS", 11)
        GeneraRubros(hoja, IdPeriodo, "INGPUB", 16, 17, False)
        GeneraRubros(hoja, IdPeriodo, "COMVOL", 19, 19, False)
        GeneraRubros(hoja, IdPeriodo, "INGVEN", 21, 21, False)

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("E.G.P. x MESES")
        For i = 1 To IdMes
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 6, 7, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 9, 9, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 11, 11, 3)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 14, 62, 3, False)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 65, 84, 3, False)
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 87, 96, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 105, 119, 3)
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("E.G.P. x MESES V2")
        For i = 1 To IdMes
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosMes(hoja, IdPeriodo, "INGPUB", IdAñoMes, 6, 7, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "COMVOL", IdAñoMes, 9, 9, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "INGVEN", IdAñoMes, 11, 11, 3)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "", IdAñoMes, 14, 62, 3, True)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "I", "", IdAñoMes, 65, 84, 3, True)
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 87, 96, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "OTRRUB", IdAñoMes, 105, 119, 3)
        Next

        IdAñoMes = IdAño * 100
        hoja = reporte.Worksheets("COSTOS x RUBROS")
        For i = 1 To IdMes
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosCostosMes(hoja, IdPeriodo, IdAñoMes, 7, 66, 3)
            GeneraRubrosMes(hoja, IdPeriodo, "OTRCOS", IdAñoMes, 69, 78, 3)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "26", IdAñoMes, 81, 81, 3, True)
            GeneraCentrosCostoMes(hoja, IdPeriodo, "D", "DEP", IdAñoMes, 84, 98, 3, True)
        Next

        GeneraCostos(reporte, IdPeriodo, "I")
        GeneraCostosAgrupados(reporte, IdPeriodo)
        GeneraCostos(reporte, IdPeriodo, "D")

        reporte.Worksheets("Plantilla Costos").Delete()
        reporte.Worksheets("COSTOS").Delete()

        hoja = reporte.Worksheets("GASTOS LEGALES")
        GeneraGastosLegDep(hoja, IdPeriodo, "GL1", 9, 48)
        GeneraGastosLegDep(hoja, IdPeriodo, "GL2", 49, 78)

        hoja = reporte.Worksheets("DEPRECIACION")
        GeneraGastosLegDep(hoja, IdPeriodo, "DEP", 8, 12)
        GeneraGastosLegDep(hoja, IdPeriodo, "AMO", 14, 14)

        hoja = reporte.Worksheets("C.V.PROD.CANJE")
        GeneraGastosBancarios(hoja, IdPeriodo, "3", 8, 38)

        hoja = reporte.Worksheets("INT.GTOS.BANC Y DEUDA L.P.")
        GeneraGastosBancarios(hoja, IdPeriodo, "2", 8, 38)

        hoja = reporte.Worksheets("INT.GTOS.BANC Y DEUDA L.P.")
        GeneraGastosBancarios(hoja, IdPeriodo, "1", 345, 375)

        hoja = reporte.Worksheets("OTROS EGRESOS")
        GeneraGastosBancarios(hoja, IdPeriodo, "4", 8, 38)

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

    Sub GeneraEGP(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal FlagAcum As Boolean, ByVal FlagVersion2 As Boolean)
        'Try

        If Not FlagAcum Then
            Titulos(hoja, IdPeriodo, "ESTADO DE GANANCIAS Y PÉRDIDAS", 4)
        Else
            TitulosComp(hoja, IdPeriodo, 4)
        End If
        GeneraRubros(hoja, IdPeriodo, "INGPUB", 9, 10, FlagAcum)
        GeneraRubros(hoja, IdPeriodo, "COMVOL", 12, 12, FlagAcum)
        GeneraRubros(hoja, IdPeriodo, "INGVEN", 14, 14, FlagAcum)
        GeneraCentrosCosto(hoja, IdPeriodo, "D", 18, 77, FlagAcum, FlagVersion2)
        GeneraCentrosCosto(hoja, IdPeriodo, "I", 80, 99, FlagAcum, FlagVersion2)
        GeneraRubros(hoja, IdPeriodo, "OTRCOS", 102, 111, FlagAcum)
        GeneraRubros(hoja, IdPeriodo, "OTRRUB", 117, 131, FlagAcum)

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraRubros(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal FlagAcum As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt, IdPeriodoAux, IdMesAux As String
        Dim array As Object(,)

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        If Not FlagAcum Then IdPeriodoAux = IdAñoAnt & "12" Else IdPeriodoAux = IdAñoAnt & IdMes
        If Not FlagAcum Then IdMesAux = "12" Else IdMesAux = IdMes

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as MontoAñoAct, isnull(C1.MontoUSD, 0) / " & IdMes & " as MontoPromAct, " & _
            "C2.MontoUSD as MontoAñoAnt, C2.MontoUSD / " & IdMesAux & " as MontoPromAnt, E.Orden " & _
            "from (select * from CtaEGP where CodSeccion ='" & CodSeccion & "') E left join " & _
            "(select IdCtaEGP, sum(MontoUSD) as MontoUSD from DetalleEGP where CodSeccion = '" & CodSeccion & "' and IdPeriodo = " & IdPeriodo & " group by IdCtaEGP) C " & _
            "on E.IdCtaEGP = C.IdCtaEGP left join " & _
            "(select IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP where CodSeccion = '" & CodSeccion & "' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaEGP) C1 " & _
            "on E.IdCtaEGP = C1.IdCtaEGP left join " & _
            "(select IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP where CodSeccion = '" & CodSeccion & "' and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAux & " group by IdCtaEGP) C2 " & _
            "on E.IdCtaEGP = C2.IdCtaEGP " & _
            "where (not C1.MontoUSD is null or not C2.MontoUSD is null) " & _
            "order by E.Orden"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        If Not FlagAcum Then
            array = DataSet2Array(dtaset.Tables(0), 2, 0, 1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dtaset.Tables(0), 4, 2, 3, 4, 5, -1)
            hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        Else
            array = DataSet2Array(dtaset.Tables(0), 5, 0, 2, 4, 3, 5)
            hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        End If

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_fin Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True
        End If
        cn.Close()
    End Sub

    Sub GeneraCentrosCosto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal FlagAcum As Boolean, ByVal FlagVersion2 As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt, IdPeriodoAux, IdMesAux As String

        Dim array As Object(,)

        'Try

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        If Not FlagAcum Then IdPeriodoAux = IdAñoAnt & "12" Else IdPeriodoAux = IdAñoAnt & IdMes
        If Not FlagAcum Then IdMesAux = "12" Else IdMesAux = IdMes

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagVersion2 Then
            sql = "select CC.CentroCostoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct, isnull(C1.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
                "C2.MontoUSD as CostoAñoAnt, C2.MontoUSD / " & IdMesAux & " as CostoPromAnt " & _
                "from (select * from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " group by CodCentroCosto) C " & _
                "on CC.CodCentroCosto = C.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) C1 " & _
                "on CC.CodCentroCosto = C1.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAux & " group by CodCentroCosto) C2 " & _
                "on CC.CodCentroCosto = C2.CodCentroCosto " & _
                "where (not C1.MontoUSD is null or not C2.MontoUSD is null) " & _
                "order by 1"
        Else
            sql = "select CC.GrupoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct, isnull(C1.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
                "C2.MontoUSD as CostoAñoAnt, C2.MontoUSD / " & IdMesAux & " as CostoPromAnt " & _
                "from (select distinct CodGrupoEGP, GrupoEGP from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " group by CodGrupoEGP) C " & _
                "on CC.CodGrupoEGP = C.CodGrupoEGP left join " & _
                "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodGrupoEGP) C1 " & _
                "on CC.CodGrupoEGP = C1.CodGrupoEGP left join " & _
                "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAux & " group by CodGrupoEGP) C2 " & _
                "on CC.CodGrupoEGP = C2.CodGrupoEGP " & _
                "where (not C1.MontoUSD is null or not C2.MontoUSD is null) " & _
                "order by 1"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        If Not FlagAcum Then
            array = DataSet2Array(dtaset.Tables(0), 2, 0, 1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dtaset.Tables(0), 4, 2, 3, 4, 5, -1)
            hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        Else
            array = DataSet2Array(dtaset.Tables(0), 5, 0, 2, 4, 3, 5)
            hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        End If

        hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraRubrosMes(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, Optional ByVal flagPpto As String = "EJEC")
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String
        Dim CampoMontoUSD As String

        Dim IdAño, IdAñoAnt As String
        Dim IdMesAux As Integer

        Dim array As Object(,)

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If flagPpto = "EJEC" Then
            CampoMontoUSD = "MontoUSD"
        ElseIf flagPpto = "PPTO" Then
            CampoMontoUSD = "PptoUSD"
            IdPeriodo = IdAño & "12"
        Else
            CampoMontoUSD = "case when IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " then MontoUSD else PptoUSD end"
            IdPeriodo = IdAño & "12"
        End If

        sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as MontoAñoAct, E.Orden " & _
            "from (select * from CtaEGP where CodSeccion ='" & CodSeccion & "') E left join " & _
            "(select IdCtaEGP, sum(" & CampoMontoUSD & ") as MontoUSD from DetalleEGP where CodSeccion = '" & CodSeccion & "' and IdPeriodo = " & IdPeriodoAux & " group by IdCtaEGP) C " & _
            "on E.IdCtaEGP = C.IdCtaEGP left join " & _
            "(select IdCtaEGP, sum(" & CampoMontoUSD & ") as MontoUSD " & _
            "from DetalleEGP where CodSeccion = '" & CodSeccion & "' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaEGP) C1 " & _
            "on E.IdCtaEGP = C1.IdCtaEGP " & _
            "where not C1.MontoUSD is null " & _
            "order by E.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 1, 0, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini)).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMesAux), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini + IdMesAux)).Value2 = array

        If fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Sub GeneraCentrosCostoMes(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String, ByVal CodCentroCosto As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, ByVal FlagVersion2 As Boolean, Optional ByVal FlagPpto As String = "EJEC")
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String
        Dim CampoMontoUSD As String

        Dim IdAño, IdAñoAnt As String
        Dim IdMesAux As Integer

        Dim array As Object(,)

        'Try
        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        'If Not FlagPpto Then CampoMontoUSD = "MontoUSD" Else CampoMontoUSD = "PptoUSD"

        If FlagPpto = "EJEC" Then
            CampoMontoUSD = "MontoUSD"
        ElseIf FlagPpto = "PPTO" Then
            CampoMontoUSD = "PptoUSD"
            IdPeriodo = IdAño & "12"
        Else
            CampoMontoUSD = "case when IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " then MontoUSD else PptoUSD end"
            IdPeriodo = IdAño & "12"
        End If

        If FlagVersion2 Then
            sql = "select CC.CentroCostoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct " & _
                "from (select * from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodCentroCosto, sum(" & CampoMontoUSD & ") as MontoUSD from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodoAux & " group by CodCentroCosto) C " & _
                "on CC.CodCentroCosto = C.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(" & CampoMontoUSD & ") as MontoUSD " & _
                "from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) C1 " & _
                "on CC.CodCentroCosto = C1.CodCentroCosto " & _
                "where not C1.MontoUSD is null "
            If CodCentroCosto = "DEP" Then
                sql = sql & "and isnull(CC.FlagDeporte, '') = 'S' "
            ElseIf CodCentroCosto <> "" Then
                sql = sql & "and CC.CodCentroCosto = " & CodCentroCosto & " "
            End If
            sql = sql & "order by 1"
        Else
            sql = "select CC.GrupoEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct " & _
                "from (select distinct CodGrupoEGP, GrupoEGP from CentroCosto where CodTipo = '" & CodTipo & "') CC left join " & _
                "(select CodGrupoEGP, sum(" & CampoMontoUSD & ") as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodoAux & " group by CodGrupoEGP) C " & _
                "on CC.CodGrupoEGP = C.CodGrupoEGP left join " & _
                "(select CodGrupoEGP, sum(" & CampoMontoUSD & ") as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodGrupoEGP) C1 " & _
                "on CC.CodGrupoEGP = C1.CodGrupoEGP " & _
                "where not C1.MontoUSD is null "
            'If CodCentroCosto = "DEP" Then
            '    sql = sql & "and isnull(CC.FlagDeporte, '') = 'S' "
            'ElseIf CodCentroCosto <> "" Then
            '    sql = sql & "and CC.CodCentroCosto = " & CodCentroCosto & " "
            'End If
            sql = sql & "order by 1"
        End If
        'sql = "select CentroCostoEGP, sum(MontoUSD) as MontoUSD " & _
        '        "from DetalleEGP E, CentroCosto C " & _
        '        "where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and C.CodTipo = '" & CodTipo & "' and IdPeriodo = " & IdPeriodo & " "
        'If CodCentroCosto = "DEP" Then
        '    sql = sql & "and isnull(FlagDeporte, '') = 'S' "
        'ElseIf CodCentroCosto <> "" Then
        '    sql = sql & "and E.CodCentroCosto = " & CodCentroCosto & " "
        'End If
        'sql = sql & "group by CentroCostoEGP order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 1, 0, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini)).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMesAux), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini + IdMesAux)).Value2 = array

        If fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraRubrosCostosMes(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, Optional ByVal FlagPpto As String = "EJEC")
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String
        Dim CampoMontoUSD As String

        Dim IdAño As String
        Dim IdMesAux As Integer
        Dim CodMatFilm As String

        Dim array As Object(,)

        IdAño = IdPeriodo.Substring(0, 4)

        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        CodMatFilm = "26"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagPpto = "EJEC" Then
            CampoMontoUSD = "MontoUSD"
        ElseIf FlagPpto = "PPTO" Then
            CampoMontoUSD = "PptoUSD"
            IdPeriodo = IdAño & "12"
        Else
            CampoMontoUSD = "case when IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " then MontoUSD else PptoUSD end"
            IdPeriodo = IdAño & "12"
        End If

        sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as MontoAñoAct " & _
            "from (select IdCtaEGP, CtaEGP from CtaEGP where CodSeccion ='COSTOS') E left join " & _
            "(select IdCtaEGP, sum(" & CampoMontoUSD & ") as MontoUSD from DetalleEGP E, CentroCosto C " & _
            "where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' " & _
            "and C.CodTipo in ('D', 'I') and isnull(FlagDeporte, '') = '' and E.CodCentroCosto <> " & CodMatFilm & " and IdPeriodo = " & IdPeriodoAux & " group by IdCtaEGP) C " & _
            "on E.IdCtaEGP = C.IdCtaEGP left join " & _
            "(select IdCtaEGP, sum(" & CampoMontoUSD & ") as MontoUSD from DetalleEGP E, CentroCosto C " & _
            "where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' " & _
            "and C.CodTipo in ('D', 'I') and isnull(FlagDeporte, '') = '' and E.CodCentroCosto <> " & CodMatFilm & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaEGP) C1 " & _
            "on E.IdCtaEGP = C1.IdCtaEGP " & _
            "where not C1.MontoUSD is null " & _
            "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 1, 0, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini)).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMesAux), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini + IdMesAux)).Value2 = array

        If fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Sub GeneraCostos(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal CodTipo As String, Optional ByVal FlagProy As Boolean = False)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String
        Dim IdAño As String
        Dim fil_ini, i As Integer
        Dim hoja As Excel.Worksheet

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not FlagProy Then
            IdAño = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        Else
            IdAño = IdPeriodo.Substring(0, 4)
        End If

        sql = "select distinct E.CodCentroCosto, GrupoEGP, CentroCostoEGP " & _
                "from DetalleEGP E, CentroCosto C where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' " & _
                "and CodTipo = '" & CodTipo & "' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "order by 2, 3"
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader

        reporte.Worksheets("COSTOS").Copy(, reporte.Worksheets("COSTOS"))
        hoja = reporte.Worksheets("COSTOS (2)")
        hoja.Name = "-"

        fil_ini = 4
        i = 0
        While rdr.Read
            Dim CentroCostoEGP = rdr("CentroCostoEGP").ToString.Replace("Costos ", "").ToUpper
            If CentroCostoEGP.Length > 15 Then CentroCostoEGP = CentroCostoEGP.Substring(0, 15)
            If hoja.Name = "-" Then
                If CodTipo = "D" Then hoja.Name = CentroCostoEGP & "-" Else hoja.Name = CentroCostoEGP
            Else
                hoja.Name = hoja.Name & CentroCostoEGP
            End If

            If Not FlagProy Then
                GeneraCostosPorCentro(reporte, hoja, rdr("CodCentroCosto"), IdPeriodo, fil_ini, False)
                fil_ini = fil_ini + 46
            Else
                GeneraCostosPorCentroEjecPpto(reporte, hoja, rdr("CodCentroCosto"), IdPeriodo, fil_ini, False)
                fil_ini = fil_ini + 49
            End If

            i = i + 1
            If (i Mod 2 = 0 Or CodTipo = "I") Then
                If Not FlagProy Then hoja.PageSetup.PrintArea = "$B$1:$I$" & fil_ini.ToString Else hoja.PageSetup.PrintArea = "$B$1:$J$" & (fil_ini - 2).ToString
                reporte.Worksheets("COSTOS").Copy(, hoja)
                hoja = reporte.Worksheets("COSTOS (2)")
                hoja.Name = "-"
                fil_ini = 4
            End If
            'If i Mod 3 = 0 Then reporte.Worksheets("COSTOS").HPageBreaks.Add(reporte.Worksheets("COSTOS").Cells(fil_ini, 1))
        End While
        If hoja.Name = "-" Then hoja.Delete()

        rdr.Close()
        cn.Close()
    End Sub

    Sub GeneraCostosAgrupados(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, Optional ByVal FlagProy As Boolean = False)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim sql As String

        Dim fil_ini, i As Integer

        Dim hoja As Excel.Worksheet

        Dim IdAñoAnt As String

        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        'Try
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodGrupoEGP, GrupoEGP " & _
                "from DetalleEGP E, CentroCosto C where E.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and CodTipo = 'I' and CodGrupoEGP in (5, 6) " & _
                "and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodo & " " & _
                "order by 2 desc"
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader

        i = 0
        While rdr.Read
            fil_ini = 4
            Dim GrupoEGP = rdr("GrupoEGP").ToString.Replace("Costos ", "").ToUpper
            reporte.Worksheets("COSTOS").Copy(, reporte.Worksheets("COSTOS"))
            hoja = reporte.Worksheets("COSTOS (2)")
            hoja.Name = GrupoEGP & " AGRUPADO"
            If Not FlagProy Then
                GeneraCostosPorCentro(reporte, hoja, rdr("CodGrupoEGP"), IdPeriodo, fil_ini, True)
            Else
                GeneraCostosPorCentroEjecPpto(reporte, hoja, rdr("CodGrupoEGP"), IdPeriodo, fil_ini, True)
            End If
            If Not FlagProy Then hoja.PageSetup.PrintArea = "$B$1:$I$" + (fil_ini + 46).ToString Else hoja.PageSetup.PrintArea = "$B$1:$J$" + (fil_ini + 45).ToString
            i = i + 1
        End While

        rdr.Close()
        cn.Close()
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub GeneraCostosPorCentro(ByVal reporte As Excel.Workbook, ByVal hoja As Excel.Worksheet, ByVal CodCentroCosto As String, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal FlagAgrupado As Boolean)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        'Dim hoja As Excel.Worksheet
        Dim rango As Excel.Range

        Dim array As Object(,)

        Dim fil As Integer
        Dim IdMes, IdAño, IdAñoAnt As String

        'Try
        'hoja = reporte.Worksheets("Plantilla Costos")
        rango = reporte.Worksheets("Plantilla Costos").Rows("4:49")
        'hoja = reporte.Worksheets("COSTOS")
        rango.Copy(hoja.Rows(fil_ini))

        Titulos(hoja, IdPeriodo, BuscaCentroCosto(CodCentroCosto).ToUpper, fil_ini)

        fil_ini = fil_ini + 4

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        fil = fil_ini
        If Not FlagAgrupado Then
            sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct, isnull(C1.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " &
                "isnull(C2.MontoUSD, 0) as CostoAñoAnt, isnull(C2.MontoUSD, 0) / 12 as CostoPromAnt " & _
                "from (select * from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                "(select CodCentroCosto, IdCtaEGP, sum(MontoUSD) as MontoUSD from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C " & _
                "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                "(select CodCentroCosto, IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C1 " & _
                "on E.IdCtaEGP = C1.IdCtaEGP left join " & _
                "(select CodCentroCosto, IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 and CodCentroCosto = " & CodCentroCosto & " group by CodCentroCosto, IdCtaEGP) C2 " & _
                "on E.IdCtaEGP = C2.IdCtaEGP " & _
                "where (not C1.MontoUSD is null or not C2.MontoUSD is null) " & _
                "order by 1"
        Else
            sql = "select E.CtaEGP, isnull(C.MontoUSD, 0) as CostoMes, isnull(C1.MontoUSD, 0) as CostoAñoAct, isnull(C1.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " &
                "isnull(C2.MontoUSD, 0) as CostoAñoAnt, isnull(C2.MontoUSD, 0) / 12 as CostoPromAnt " & _
                "from (select * from CtaEGP where CodSeccion = 'COSTOS') E left join " & _
                "(select IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo = " & IdPeriodo & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP ) C " & _
                "on E.IdCtaEGP = C.IdCtaEGP left join " & _
                "(select IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP) C1 " & _
                "on E.IdCtaEGP = C1.IdCtaEGP left join " & _
                "(select IdCtaEGP, sum(MontoUSD) as MontoUSD " & _
                "from DetalleEGP D, CentroCosto C where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 and CodGrupoEGP = " & CodCentroCosto & " group by IdCtaEGP) C2 " & _
                "on E.IdCtaEGP = C2.IdCtaEGP " & _
                "where (not C1.MontoUSD is null or not C2.MontoUSD is null) " & _
                "order by 1"
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 2, 0, 1, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 4, 2, 3, 4, 5, -1)
        hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_ini + 40 - 1 Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & (fil_ini + 40 - 1).ToString).EntireRow.Hidden = True
        End If

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Sub GeneraGastosBancarios(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dr As DataRow()
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt As String
        Dim array As Object(,)

        Titulos(hoja, IdPeriodo, "", 4)

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select *, ROW_NUMBER() over (order by AcumAct desc, AcumAnt desc, Persona) as Orden from " & _
                "(select isnull(G1.Persona, G2.Persona) as Persona, isnull(G.MontoUSD, 0) as MontoMes, " & _
                "isnull(G1.MontoUSD, 0) as AcumAct, isnull(G1.MontoUSD, 0) / " & IdMes & " as PromAct, isnull(G2.MontoUSD, 0) as AcumAnt, isnull(G2.MontoUSD, 0) / 12 as PromAnt " & _
                "from (select CodTipo, CodPersona, Persona, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoBan where IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodTipo, CodPersona, Persona) G1 full join " & _
                "(select CodTipo, CodPersona, Persona, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoBan where IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 " & _
                "group by CodTipo, CodPersona, Persona) G2 " & _
                "on G1.CodTipo = G2.CodTipo and G1.CodPersona = G2.CodPersona and G1.Persona = G2.Persona left join " & _
                "(select CodTipo, CodPersona, Persona, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoBan where IdPeriodo = " & IdPeriodo & " group by CodTipo, CodPersona, Persona) G " & _
                "on G1.CodTipo = G.CodTipo and G1.CodPersona = G.CodPersona and G1.Persona = G.Persona " & _
                "where (G1.CodTipo = " & CodTipo & " or G2.CodTipo = " & CodTipo & ")) as T "
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        dr = dtaset.Tables(0).Select("Orden <= 30", "Persona")
        If dr.GetUpperBound(0) > 0 Then
            array = DataSet2Array(dr, 2, 0, 1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dr, 4, 2, 3, 4, 5, -1)
            hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array
        End If

        fil_ini = fil_ini + 41
        dr = dtaset.Tables(0).Select("Orden > 30", "Persona")
        If dr.GetUpperBound(0) > 0 Then
            array = DataSet2Array(dr, 2, 0, 1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dr, 4, 2, 3, 4, 5, -1)
            hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array
        End If

        'If fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Sub GeneraGastosLegDep(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt As String
        Dim array As Object(,)

        Titulos(hoja, IdPeriodo, "", 4)

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(G1.Rubro, G2.Rubro) as Rubro, isnull(G.MontoUSD, 0) as MontoMes, " & _
                "isnull(G1.MontoUSD, 0) as AcumAct, isnull(G1.MontoUSD, 0) / " & IdMes & " as PromAct, isnull(G2.MontoUSD, 0) as AcumAnt, isnull(G2.MontoUSD, 0) / 12 as PromAnt from " & _
                "(select Rubro, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoLegal where CodTipo = '" & CodTipo & "' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Rubro having sum(MontoUSD) <> 0) G1 full join " & _
                "(select Rubro, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoLegal where CodTipo = '" & CodTipo & "' and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 " & _
                "group by Rubro having sum(MontoUSD) <> 0) G2 " & _
                "on G1.Rubro = G2.Rubro left join " & _
                "(select Rubro, sum(MontoUSD) as MontoUSD " & _
                "from EGPGastoLegal where CodTipo = '" & CodTipo & "' and IdPeriodo = " & IdPeriodo & " " & _
                "group by Rubro having sum(MontoUSD) <> 0) G " & _
                "on G1.Rubro = G.Rubro order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        If dtaset.Tables(0).Rows.Count > 0 Then
            array = DataSet2Array(dtaset.Tables(0), 2, 0, 1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dtaset.Tables(0), 4, 2, 3, 4, 5, -1)
            hoja.Range("E" & fil_ini.ToString & ":H" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array
        End If

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_fin Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraReportesComparativos(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String, ByVal CodCentroCosto As String, ByVal IdCtaEGP As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        Dim IdAñoAnt As Integer

        'Try

        IdAñoAnt = Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("COMPARATIVO CUENTAS")
        hoja.Cells(4, 2).Value = BuscaCentroCosto(CodCentroCosto).ToUpper & " - " & BuscaRubroCosto(IdCtaEGP).ToUpper
        hoja.Cells(5, 2).Value = "AL MES DE " & BuscaPeriodo(IdPeriodo).ToUpper & " - " & IdAñoAnt.ToString

        If IdAñoAnt >= 2011 Then
            hoja.Cells(7, 2).Value = "Cuenta"
            hoja.Columns(3).Hidden = True
        Else
            hoja.Cells(7, 2).Value = "Cuenta " & IdPeriodo.Substring(0, 4)
            hoja.Cells(7, 3).Value = "Cuenta " & IdAñoAnt.ToString
        End If

        hoja.Cells(7, 5).Value = "ACUMULADO " & IdPeriodo.Substring(0, 4)
        hoja.Cells(7, 6).Value = "ACUMULADO " & IdAñoAnt.ToString

        GeneraComparativoCuentas(hoja, IdPeriodo, CodCentroCosto, IdCtaEGP)

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

    Sub GeneraComparativoCuentas(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodCentroCosto As String, ByVal IdCtaEGP As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim array As Object(,)

        Dim IdAño, IdAñoAnt As Integer
        Dim IdPeriodoAnt As String

        Dim fil_ini As Integer

        'Try
        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdAñoAnt = IdAño - 1
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        fil_ini = 8

        If Convert.ToInt32(CodCentroCosto) < 1000 Then
            If IdAñoAnt >= 2011 Then
                sql = "select isnull(A.CodCuenta, A1.CodCuenta) as Codcuenta, null as CodCuentaAnt, isnull(A.Cuenta, A1.Cuenta) as Cuenta, isnull(AcumAct, 0) as AcumAct, isnull(AcumAnt, 0) as AcumAnt " & _
                        "from " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAct " & _
                        "from Asiento A, CtaEGPDet D " & _
                        "where A.CodCuenta = D.CodCtaOrigen and CodCentroCosto = " & CodCentroCosto & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAño.ToString & "01 and " & IdPeriodo & " " & _
                        "group by CodCuenta, Cuenta) A full join " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAnt " & _
                        "from Asiento A, CtaEGPDet D " & _
                        "where A.CodCuenta = D.CodCtaOrigen and CodCentroCosto = " & CodCentroCosto & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt.ToString & "01 and " & IdPeriodoAnt & " " & _
                        "group by CodCuenta, Cuenta) A1 on A.CodCuenta = A1.CodCuenta"
            Else
                sql = "select isnull(A.CodCuenta, A1.CodCuenta) as Codcuenta, CodCuentaAnt, isnull(A.Cuenta, A1.Cuenta) as Cuenta, isnull(AcumAct, 0) as AcumAct, isnull(AcumAnt, 0) as AcumAnt " & _
                        "from " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAct " & _
                        "from Asiento A, CtaEGPDet D " & _
                        "where A.CodCuenta = D.CodCtaOrigen and CodCentroCosto = " & CodCentroCosto & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAño.ToString & "01 and " & IdPeriodo & " " & _
                        "group by CodCuenta, Cuenta) A full join " & _
                        "(select E.CodCuenta, C.Cuenta, min(E.CodCuentaAnt) as CodCuentaAnt, sum(DebeUSD) - sum(HaberUSD) as AcumAnt " & _
                        "from Asiento A, V_EquivCuentaEGP E, CuentaEGP C, CtaEGPDet D " & _
                        "where A.CodCuenta = E.CodCuentaAnt and E.CodCuenta = C.CodCuenta and E.CodCuenta = D.CodCtaOrigen and CodCentroCosto = " & CodCentroCosto & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt.ToString & "01 and " & IdPeriodoAnt & " " & _
                        "group by E.CodCuenta, C.Cuenta) A1 on A.CodCuenta = A1.CodCuenta"
            End If
        Else
            Dim CodGrupoEGP As String = (Convert.ToInt32(CodCentroCosto) - 1000).ToString
            If IdAñoAnt >= 2011 Then
                sql = "select isnull(A.CodCuenta, A1.CodCuenta) as Codcuenta, null as CodCuentaAnt, isnull(A.Cuenta, A1.Cuenta) as Cuenta, isnull(AcumAct, 0) as AcumAct, isnull(AcumAnt, 0) as AcumAnt " & _
                        "from " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAct " & _
                        "from Asiento A, CtaEGPDet D, CentroCosto CC " & _
                        "where A.CodCuenta = D.CodCtaOrigen and A.CodCentroCosto = CC.CodCentroCosto and CodGrupoEGP = " & CodGrupoEGP & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAño.ToString & "01 and " & IdPeriodo & " " & _
                        "group by CodCuenta, Cuenta) A full join " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAnt " & _
                        "from Asiento A, CtaEGPDet D, CentroCosto CC " & _
                        "where A.CodCuenta = D.CodCtaOrigen and A.CodCentroCosto = CC.CodCentroCosto and CodGrupoEGP = " & CodGrupoEGP & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt.ToString & "01 and " & IdPeriodoAnt & " " & _
                        "group by CodCuenta, Cuenta) A1 on A.CodCuenta = A1.CodCuenta"
            Else
                sql = "select isnull(A.CodCuenta, A1.CodCuenta) as Codcuenta, CodCuentaAnt, isnull(A.Cuenta, A1.Cuenta) as Cuenta, isnull(AcumAct, 0) as AcumAct, isnull(AcumAnt, 0) as AcumAnt " & _
                        "from " & _
                        "(select CodCuenta, Cuenta, sum(DebeUSD) - sum(HaberUSD) as AcumAct " & _
                        "from Asiento A, CtaEGPDet D, CentroCosto CC " & _
                        "where A.CodCuenta = D.CodCtaOrigen and A.CodCentroCosto = CC.CodCentroCosto and CodGrupoEGP = " & CodGrupoEGP & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAño.ToString & "01 and " & IdPeriodo & " " & _
                        "group by CodCuenta, Cuenta) A full join " & _
                        "(select E.CodCuenta, C.Cuenta, min(E.CodCuentaAnt) as CodCuentaAnt, sum(DebeUSD) - sum(HaberUSD) as AcumAnt " & _
                        "from Asiento A, V_EquivCuentaEGP E, CuentaEGP C, CtaEGPDet D, CentroCosto CC " & _
                        "where A.CodCuenta = E.CodCuentaAnt and E.CodCuenta = C.CodCuenta and E.CodCuenta = D.CodCtaOrigen and A.CodCentroCosto = CC.CodCentroCosto " & _
                        "and CodGrupoEGP = " & CodGrupoEGP & " and D.IdCtaEGP = " & IdCtaEGP & " " & _
                        "and year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt.ToString & "01 and " & IdPeriodoAnt & " " & _
                        "group by E.CodCuenta, C.Cuenta) A1 on A.CodCuenta = A1.CodCuenta"
            End If
        End If

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), False)
        hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + dtaset.Tables(0).Rows.Count - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + dtaset.Tables(0).Rows.Count <= fil_ini + 80 - 1 Then
            hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & (fil_ini + 80 - 1).ToString).EntireRow.Hidden = True
        End If

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

End Module

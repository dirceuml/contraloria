Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesMaterialFilmico

    Function LlenaMaterialFilmico(ByVal MaterialB As String, ByVal NumContratoB As String, ByVal Orden As String, ByVal TipoOrden As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select M.*, case when isnull(CantCapitulos, 0) <> 0 then isnull(MontoMaterialUSD, 0) / CantCapitulos else isnull(MontoMaterialUSD, 0) end as CostoCapituloUSD " & _
                "from MaterialFilmico M " & _
                "where Material like '%" & MaterialB & "%' and NumContrato like '%" & NumContratoB & "%' " & _
                "and CodMaterial in (select distinct CodMaterial from ProgMaterialFilmico) " & _
                "order by " & Orden & " " & TipoOrden

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaConsumoMaterialFilmico(ByVal CodMaterial As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdConsumoMaterial, Periodo, CantCapitulosProg, MontoUSD " & _
                "from ConsumoMaterialFilmico C, Periodo P " & _
                "where C.IdPeriodo = P.IdPeriodo and CodMaterial = " & CodMaterial & " " & _
                "order by C.IdPeriodo"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaProgMaterialFilmico(ByVal CodMaterial As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdProgMaterial, Periodo, FechaProg, Programa, NumCapitulo, Rating " & _
                "from ProgMaterialFilmico P, Periodo PE " & _
                "where year(FechaProg) * 100 + month(FechaProg) = IdPeriodo and CodMaterial = " & CodMaterial & " " & _
                "order by IdPeriodo, FechaProg"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function BuscaMaterialFilmico(ByVal CodMaterial As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim Material As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Material from MaterialFilmico where CodMaterial = " & CodMaterial
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Material = rdr("Material").ToString()

        rdr.Close()
        cn.Close()

        Return Material
    End Function

    Function LlenaNumContrato() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select '' as NumContrato, '(Todos)' as NumContrato2 " & _
                "union select distinct NumContrato, NumContrato as NumContrato2 from MaterialFilmico " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub GeneraReportesMaterialFilmico(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet

        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("CERCEDILLA")
        GeneraContratos(hoja, IdPeriodo)

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

    Private Sub GeneraContratos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim UltDiaMes As Date

        Dim cant, cant2 As Integer

        Dim fil_ini As Integer

        Dim NumContrato As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select NumContrato, count(*) as Cant " & _
                "from V_MaterialFilmico M " & _
                "where NumContrato like '2009%' or NumContrato like '2010%' or NumContrato like '2011%' " & _
                "group by NumContrato order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 1, 4).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        fil_ini = fil_ini + 6

        For i = 0 To cant - 1
            NumContrato = dt.Rows(i)("NumContrato")
            cant2 = Convert.ToInt32(dt.Rows(i)("Cant"))

            hoja.Cells(fil_ini, 4).Value = NumContrato
            fil_ini = fil_ini + 1

            'If cant2 > 10 Then
            '    Dim rango As Excel.Range
            '    rango = hoja.Rows((fil_ini + 1).ToString & ":" & (fil_ini + cant - 10).ToString)
            '    'rango.Insert(, (hoja.Rows(fil_ini))
            '    rango.Insert(, rango)
            'End If

            For j = 1 To 12
                GeneraContrato(hoja, IdPeriodo, j, NumContrato, fil_ini)
            Next

            fil_ini = fil_ini + 31
            'If cant2 <= 10 Then fil_ini = fil_ini + 11 Else fil_ini = fil_ini + cant2 + 1

            If i = 7 Or i = 15 Or i = 23 Or i = 31 Then
                'hoja.HPageBreaks.Add(hoja.Rows(fil_ini + 4))
                fil_ini = fil_ini + 10
            End If

        Next

        'hoja.PageSetup.PrintArea = "$A$1:$U$" + CStr(fil_ini - 1)

        hoja.Rows("1263:1326").EntireRow.Hidden = True


        cn.Close()
    End Sub

    Private Sub GeneraContrato(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdMes As Integer, ByVal NumContrato As String, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdPeriodoMes As String

        Dim array As Object(,)
        Dim cant As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoMes = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If IdMes = 1 Then
            sql = "select M.Material, Genero, Contrato, Contrato - isnull(ConsumoAnt, 0) as SaldoAnt " & _
                    "from V_MaterialFilmico M left join " & _
                    "(select NumContrato, Material, sum(Consumo) as ConsumoAnt " & _
                    "from V_ConsumoMaterialFilmico " & _
                    "where IdPeriodo < " & IdAño & "01 " & _
                    "group by NumContrato, Material) C on M.NumContrato = C.NumContrato and M.Material = C.Material " & _
                    "where M.Numcontrato = '" & NumContrato & "' order by Genero, Material"

            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)
            cant = dt.Rows.Count

            array = DataSet2Array(dt, 4, 0, 1, 2, 3, -1)
            hoja.Range("D" & fil_ini.ToString & ":G" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        End If

        sql = "select M.Material, isnull(Consumo, 0) as Consumo " & _
                "from V_MaterialFilmico M left join " & _
                "(select NumContrato, Material, sum(Consumo) as Consumo " & _
                "from V_ConsumoMaterialFilmico " & _
                "where IdPeriodo = " & IdPeriodoMes & " " & _
                "group by NumContrato, Material) C on M.NumContrato = C.NumContrato and M.Material = C.Material " & _
                "where M.NumContrato = '" & NumContrato & "' order by Genero, Material"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range(Chr(65 + 7 + IdMes - 1) & fil_ini.ToString & ":" & Chr(65 + 7 + IdMes - 1) & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < 30 Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + 30 - 1).ToString).EntireRow.Hidden = True
        'If IdMes = 12 Then
        '    fil_ini = fil_ini + cant
        'End If

        cn.Close()
    End Sub

End Module

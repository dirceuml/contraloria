Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesRentabilidad

    Sub DistribucionCostos(ByVal Peso As Double)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "sp_Distribucion_Costos"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Peso", Peso)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub GeneraReportesRentabilidad(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String, ByVal Peso As Double)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet

        'Try

        FuncionesEGP.CreaDetalleEGP(IdPeriodo)
        DistribucionCostos(Peso / 100)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("RESUMEN")
        GeneraResumen(hoja, IdPeriodo)
        hoja = reporte.Worksheets("RESUMEN COMPARATIVO")
        GeneraResumenComparativo(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COMPARATIVO DETALLADO")
        GeneraComparativoDetallado(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COMPARATIVO C. DIRECTOS")
        GeneraComparativoCostosDirectos(hoja, IdPeriodo)
        hoja = reporte.Worksheets("C. DIRECTOS")
        GeneraCostosAgrupados(hoja, IdPeriodo, "D")
        hoja = reporte.Worksheets("C. INDIRECTOS")
        GeneraCostosAgrupados(hoja, IdPeriodo, "I")
        hoja = reporte.Worksheets("C. INDIRECTOS DETALLADO")
        GeneraCostosIndirectos(hoja, IdPeriodo)

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

    Private Sub GeneraResumen(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        CantReg = 50

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(A.Programa, A1.Programa) as Programa, isnull(CostoDirecto, 0) as CostoDirecto, isnull(CostoIndirecto, 0) as CostoIndirecto, " & _
                "isnull(Venta, 0) as Venta " & _
                "from " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoDirecto) as CostoDirecto, sum(CostoIndirecto) as CostoIndirecto " & _
                "from Distribucion where FlagMaterial is null and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by upper(GrupoPrograma)) A full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as Venta " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo union " & _
                "select 'Material Filmico', isnull(sum(MontoUSD), 0) " & _
                "from V_Facturacion where CodGrupo = 4 and FlagMaterial = 'S' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ") A1 on A.Programa = A1.Programa " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        fil_ini = fil_ini + 7

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraResumenComparativo(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt, IdPeriodoCargaAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        CantReg = 50

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)
        IdPeriodoCargaAnt = IdAñoAnt & "12"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(A.Programa, isnull(A1.Programa, isnull(A2.Programa, A3.Programa))) as Programa, isnull(CostoTotalAct, 0) as CostoTotalAct, " & _
                "isnull(CostoTotalAnt, 0) as CostoTotalAnt, isnull(VentaAct, 0) as VentaAct, isnull(VentaAnt, 0) as VentaAnt " & _
                "from " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoTotal) as CostoTotalAct " & _
                "from Distribucion where GrupoPrograma <> 'Material Filmico' and FlagMaterial is null and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by upper(GrupoPrograma)) A full join " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoTotal) as CostoTotalAnt " & _
                "from Distribucion where GrupoPrograma <> 'Material Filmico' and FlagMaterial is null and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by upper(GrupoPrograma)) A1 on A.Programa = A1.Programa full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAct " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo) A2 on (A.Programa = A2.Programa or A1.Programa = A2.Programa) full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAnt " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodoCargaAnt & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by Grupo) A3 on (A.Programa = A3.Programa or A1.Programa = A3.Programa or A2.Programa = A3.Programa) " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 3).Value = IdAño
        hoja.Cells(fil_ini + 7, 4).Value = IdAñoAnt
        hoja.Cells(fil_ini + 7, 5).Value = IdAño
        hoja.Cells(fil_ini + 7, 6).Value = IdAñoAnt
        hoja.Cells(fil_ini + 7, 7).Value = IdAño
        hoja.Cells(fil_ini + 7, 8).Value = IdAñoAnt

        fil_ini = fil_ini + 8

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        sql = "select isnull(A.Programa, isnull(A1.Programa, isnull(A2.Programa, A3.Programa))) as Programa, isnull(CostoTotalAct, 0) as CostoTotalAct, " & _
                "isnull(CostoTotalAnt, 0) as CostoTotalAnt, isnull(VentaAct, 0) as VentaAct, isnull(VentaAnt, 0) as VentaAnt " & _
                "from " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoTotal) as CostoTotalAct " & _
                "from Distribucion where FlagMaterial = 'S' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by upper(GrupoPrograma)) A full join " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoTotal) as CostoTotalAnt " & _
                "from Distribucion where FlagMaterial = 'S' and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by upper(GrupoPrograma)) A1 on A.Programa = A1.Programa full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAct " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial = 'S' and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo) A2 on (A.Programa = A2.Programa or A1.Programa = A2.Programa) full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAnt " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial = 'S' and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodoCargaAnt & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by Grupo) A3 on (A.Programa = A3.Programa or A1.Programa = A3.Programa or A2.Programa = A3.Programa) " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        CantReg = 60
        fil_ini = 61
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 3).Value = IdAño
        hoja.Cells(fil_ini + 7, 4).Value = IdAñoAnt
        hoja.Cells(fil_ini + 7, 5).Value = IdAño
        hoja.Cells(fil_ini + 7, 6).Value = IdAñoAnt
        hoja.Cells(fil_ini + 7, 7).Value = IdAño
        hoja.Cells(fil_ini + 7, 8).Value = IdAñoAnt

        fil_ini = fil_ini + 8

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraComparativoDetallado(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt, IdPeriodoCargaAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        CantReg = 50

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)
        IdPeriodoCargaAnt = IdAñoAnt & "12"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(A.Programa, isnull(A1.Programa, isnull(A2.Programa, A3.Programa))) as Programa, " & _
                "isnull(CostoDirectoAct, 0) as CostoDirectoAct, isnull(CostoDirectoAnt, 0) as CostoDirectoAnt," & _
                "isnull(CostoIndirectoAct, 0) as CostoIndirectoAct, isnull(CostoIndirectoAnt, 0) as CostoIndirectoAnt," & _
                "isnull(VentaAct, 0) as VentaAct, isnull(VentaAnt, 0) as VentaAnt " & _
                "from " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoDirecto) as CostoDirectoAct, sum(CostoIndirecto) as CostoIndirectoAct " & _
                "from Distribucion where GrupoPrograma <> 'Material Filmico' and FlagMaterial is null and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by upper(GrupoPrograma)) A full join " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoDirecto) as CostoDirectoAnt, sum(CostoIndirecto) as CostoIndirectoAnt " & _
                "from Distribucion where GrupoPrograma <> 'Material Filmico' and FlagMaterial is null and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by upper(GrupoPrograma)) A1 on A.Programa = A1.Programa full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAct " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo) A2 on (A.Programa = A2.Programa or A1.Programa = A2.Programa) full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAnt " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodoCargaAnt & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by Grupo) A3 on (A.Programa = A3.Programa or A1.Programa = A3.Programa or A2.Programa = A3.Programa) " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 1, 4).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 58, 4).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        hoja.Cells(fil_ini + 4, 5).Value = IdAño
        hoja.Cells(fil_ini + 4, 6).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 7).Value = IdAño
        hoja.Cells(fil_ini + 4, 8).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 9).Value = IdAño
        hoja.Cells(fil_ini + 4, 10).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 11).Value = IdAño
        hoja.Cells(fil_ini + 4, 12).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 13).Value = IdAño
        hoja.Cells(fil_ini + 4, 14).Value = IdAñoAnt

        fil_ini = fil_ini + 5

        array = DataSet2Array(dt, 5, 0, 1, 2, 3, 4)
        hoja.Range("D" & fil_ini.ToString & ":H" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 5, 6, -1, -1, -1)
        hoja.Range("K" & fil_ini.ToString & ":L" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        sql = "select isnull(A.Programa, isnull(A1.Programa, isnull(A2.Programa, A3.Programa))) as Programa, " & _
                "isnull(CostoDirectoAct, 0) as CostoDirectoAct, isnull(CostoDirectoAnt, 0) as CostoDirectoAnt," & _
                "isnull(CostoIndirectoAct, 0) as CostoIndirectoAct, isnull(CostoIndirectoAnt, 0) as CostoIndirectoAnt," & _
                "isnull(VentaAct, 0) as VentaAct, isnull(VentaAnt, 0) as VentaAnt " & _
                "from " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoDirecto) as CostoDirectoAct, sum(CostoIndirecto) as CostoIndirectoAct " & _
                "from Distribucion where FlagMaterial = 'S' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by upper(GrupoPrograma)) A full join " & _
                "(select upper(GrupoPrograma) as Programa, sum(CostoDirecto) as CostoDirectoAnt, sum(CostoIndirecto) as CostoIndirectoAnt " & _
                "from Distribucion where FlagMaterial = 'S' and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by upper(GrupoPrograma)) A1 on A.Programa = A1.Programa full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAct " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial = 'S' and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo) A2 on (A.Programa = A2.Programa or A1.Programa = A2.Programa) full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAnt " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial = 'S' and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodoCargaAnt & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by Grupo) A3 on (A.Programa = A3.Programa or A1.Programa = A3.Programa or A2.Programa = A3.Programa) " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        CantReg = 60
        fil_ini = 58
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 1, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 4, 5).Value = IdAño
        hoja.Cells(fil_ini + 4, 6).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 7).Value = IdAño
        hoja.Cells(fil_ini + 4, 8).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 9).Value = IdAño
        hoja.Cells(fil_ini + 4, 10).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 11).Value = IdAño
        hoja.Cells(fil_ini + 4, 12).Value = IdAñoAnt
        hoja.Cells(fil_ini + 4, 13).Value = IdAño
        hoja.Cells(fil_ini + 4, 14).Value = IdAñoAnt

        fil_ini = fil_ini + 5

        array = DataSet2Array(dt, 5, 0, 1, 2, 3, 4)
        hoja.Range("D" & fil_ini.ToString & ":H" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 5, 6, -1, -1, -1)
        hoja.Range("K" & fil_ini.ToString & ":L" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraComparativoCostosDirectos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt, IdPeriodoCargaAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        CantReg = 50

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)
        IdPeriodoCargaAnt = IdAñoAnt & "12"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select isnull(A.Programa, isnull(A1.Programa, isnull(A2.Programa, A3.Programa))) as Programa, isnull(CostoAcumAct, 0) as CostoAcumAct, " & _
                "isnull(CostoAcumAnt, 0) as CostoAcumAnt, isnull(VentaAcumAct, 0) as VentaAcumAct, isnull(VentaAcumAnt, 0) as VentaAcumAnt " & _
                "from " & _
                "(select upper(replace(C.GrupoEGP, 'Costos ', '')) as Programa, isnull(sum(MontoUSD), 0) as CostoAcumAct from " & _
                "DetalleEGP D, CentroCosto C " & _
                "where D.CodCentroCosto = C.CodCentroCosto and D.CodSeccion = 'COSTOS' and C.CodTipo = 'D' " & _
                "and C.CodCentroCosto <> 26 and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by replace(C.GrupoEGP, 'Costos ', '')) A full join " & _
                "(select upper(replace(C.GrupoEGP, 'Costos ', '')) as Programa, isnull(sum(MontoUSD), 0) as CostoAcumAnt from " & _
                "DetalleEGP D, CentroCosto C " & _
                "where D.CodCentroCosto = C.CodCentroCosto and D.CodSeccion = 'COSTOS' and C.CodTipo = 'D' " & _
                "and C.CodCentroCosto <> 26 and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by replace(C.GrupoEGP, 'Costos ', '')) A1 on A.Programa = A1.Programa full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAcumAct " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo) A2 on (A.Programa = A2.Programa or A1.Programa = A2.Programa) full join " & _
                "(select Grupo as Programa, isnull(sum(MontoUSD), 0) as VentaAcumAnt " & _
                "from V_Facturacion F " & _
                "where CodGrupo = 4 and FlagMaterial is null and Grupo not like '*%' and Grupo <> 'SIN PROGRAMA' " & _
                "and IdPeriodoCarga = " & IdPeriodoCargaAnt & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by Grupo) A3 on (A.Programa = A3.Programa or A1.Programa = A3.Programa or A2.Programa = A3.Programa) " & _
                "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 6, 3).Value = "COSTOS A " & BuscaMes(IdPeriodo.Substring(4, 2)).ToUpper
        hoja.Cells(fil_ini + 6, 7).Value = "VENTAS A " & BuscaMes(IdPeriodo.Substring(4, 2)).ToUpper
        hoja.Cells(fil_ini + 8, 3).Value = IdAño
        hoja.Cells(fil_ini + 8, 4).Value = IdAñoAnt
        hoja.Cells(fil_ini + 8, 7).Value = IdAño
        hoja.Cells(fil_ini + 8, 8).Value = IdAñoAnt

        fil_ini = fil_ini + 9

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":D" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":H" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraCostosAgrupados(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        'Try
        If CodTipo = "D" Then CantReg = 50 Else CantReg = 20

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select upper(replace(CC.GrupoEGP, 'Costos ', '')), isnull(C.MontoUSD, 0) as CostoAñoAct, isnull(C.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
            "isnull(C1.MontoUSD, 0) as CostoAñoAnt, isnull(C1.MontoUSD, 0) / 12 as CostoPromAnt " & _
            "from (select distinct CodGrupoEGP, GrupoEGP from CentroCosto where CodTipo = '" & CodTipo & "' and CodCentroCosto <> 26) CC left join " & _
            "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP D, CentroCosto C " & _
            "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodGrupoEGP) C " & _
            "on CC.CodGrupoEGP = C.CodGrupoEGP left join " & _
            "(select CodGrupoEGP, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP D, CentroCosto C " & _
            "where D.CodCentroCosto = C.CodCentroCosto and CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 group by CodGrupoEGP) C1 " & _
            "on CC.CodGrupoEGP = C1.CodGrupoEGP " & _
            "where (not C.MontoUSD is null or not C1.MontoUSD is null) " & _
            "order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 6, 3).Value = "A " & BuscaMes(IdMes).ToUpper & " " & IdAño
        hoja.Cells(fil_ini + 6, 4).Value = "PROMEDIO " & IdAño
        hoja.Cells(fil_ini + 6, 5).Value = "A DICIEMBRE " & IdAñoAnt
        hoja.Cells(fil_ini + 6, 6).Value = "PROMEDIO " & IdAñoAnt

        fil_ini = fil_ini + 7

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":F" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        If CodTipo = "I" Then
            sql = "select 'Depreciación y Amortización' as Rubro, isnull(C.MontoUSD, 0) as CostoAñoAct, isnull(C.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
                    "isnull(C1.MontoUSD, 0) as CostoAñoAnt, isnull(C1.MontoUSD, 0) / 12 as CostoPromAnt " & _
                    "from " & _
                    "(select sum(MontoUSD) as MontoUSD " & _
                    "from DetalleEGP where IdCtaEGP = 45 and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ") C, " & _
                    "(select sum(MontoUSD) as MontoUSD " & _
                    "from DetalleEGP where IdCtaEGP = 45 and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12) C1 "
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)
            array = DataSet2Array(dt, False)
            hoja.Range("B30:F30", Type.Missing).Value2 = array
        End If
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

    Private Sub GeneraCostosIndirectos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdMes, IdAño, IdAñoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, CantReg As Integer

        Dim fil_ini As Integer

        'Try
        CantReg = 20

        IdAño = IdPeriodo.Substring(0, 4)
        IdMes = IdPeriodo.Substring(4, 2)
        IdAñoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select upper(replace(CC.GrupoEGP, 'Costos ', '')), upper(replace(CC.CentroCostoEGP, 'Costos ', '')), isnull(C.MontoUSD, 0) as CostoAñoAct, isnull(C.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
            "isnull(C1.MontoUSD, 0) as CostoAñoAnt, isnull(C1.MontoUSD, 0) / 12 as CostoPromAnt " & _
            "from (select * from CentroCosto where CodTipo = 'I') CC left join " & _
            "(select CodCentroCosto, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP D " & _
            "where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) C " & _
            "on CC.CodCentroCosto = C.CodCentroCosto left join " & _
            "(select CodCentroCosto, sum(MontoUSD) as MontoUSD " & _
            "from DetalleEGP D " & _
            "where CodSeccion = 'COSTOS' and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12 group by CodCentroCosto) C1 " & _
            "on CC.CodCentroCosto = C1.CodCentroCosto " & _
            "where (not C.MontoUSD is null or not C1.MontoUSD is null) " & _
            "order by 1, 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 6, 4).Value = "A " & BuscaMes(IdMes).ToUpper & " " & IdAño
        hoja.Cells(fil_ini + 6, 5).Value = "PROMEDIO " & IdAño
        hoja.Cells(fil_ini + 6, 6).Value = "A DICIEMBRE " & IdAñoAnt
        hoja.Cells(fil_ini + 6, 7).Value = "PROMEDIO " & IdAñoAnt

        fil_ini = fil_ini + 7

        array = DataSet2Array(dt, False)
        hoja.Range("B" & fil_ini.ToString & ":G" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < CantReg Then hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True

        sql = "select 'Depreciación y Amortización' as Rubro, null as Rubro2, isnull(C.MontoUSD, 0) as CostoAñoAct, isnull(C.MontoUSD, 0) / " & IdMes & " as CostoPromAct, " & _
        "isnull(C1.MontoUSD, 0) as CostoAñoAnt, isnull(C1.MontoUSD, 0) / 12 as CostoPromAnt " & _
        "from " & _
        "(select sum(MontoUSD) as MontoUSD " & _
        "from DetalleEGP where IdCtaEGP = 45 and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ") C, " & _
        "(select sum(MontoUSD) as MontoUSD " & _
        "from DetalleEGP where IdCtaEGP = 45 and IdPeriodo between " & IdAñoAnt & "01 and " & IdAñoAnt & "12) C1 "
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        array = DataSet2Array(dt, False)
        hoja.Range("B30:G30", Type.Missing).Value2 = array

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        '    cn.Close()
        'End Try
    End Sub

End Module

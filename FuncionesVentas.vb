Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesVentas

    Function LlenaProgramas(ByVal ProgramaB As String, ByVal Orden As String, ByVal TipoOrden As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPrograma, Programa, Genero, GrupoPrograma, isnull(FlagMaterial, 'N') as FlagMaterial " & _
                "from Programa " & _
                "where Programa like '%" & ProgramaB & "%' " & _
                "order by " & Orden & " " & TipoOrden

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaGruposPrograma() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select '(SIN GRUPO)' as GrupoPrograma " & _
                "union select distinct GrupoPrograma from Programa where GrupoPrograma is not null " & _
                "union select distinct GrupoPrograma from MaterialFilmico where GrupoPrograma is not null " & _
                "and GrupoPrograma not in (select distinct GrupoPrograma from Programa where GrupoPrograma is not null) " & _
                "order by GrupoPrograma"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaGruposProgramaMaterial() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select '(SIN GRUPO)' as GrupoPrograma " & _
                "union select distinct GrupoPrograma from Programa where FlagMaterial = 'S' and GrupoPrograma is not null " & _
                "union select distinct GrupoPrograma from MaterialFilmico where GrupoPrograma is not null " & _
                "and GrupoPrograma not in (select distinct GrupoPrograma from Programa where GrupoPrograma is not null) " & _
                "order by GrupoPrograma"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub ActualizaGrupoPrograma(ByVal IdPrograma As String, ByVal GrupoPrograma As String, ByVal FlagMaterial As Boolean)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim FlagMaterial2 As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If GrupoPrograma <> "" Then GrupoPrograma = "'" & GrupoPrograma & "'" Else GrupoPrograma = "null"
        If FlagMaterial Then FlagMaterial2 = "'S'" Else FlagMaterial2 = "null"
        sql = "update Programa set GrupoPrograma = " & GrupoPrograma & ", FlagMaterial = " & FlagMaterial2 & _
                " where IdPrograma = " & IdPrograma
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaGrupoPrograma(ByVal IdMaterial As String, ByVal GrupoPrograma As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If GrupoPrograma <> "" Then GrupoPrograma = "'" & GrupoPrograma & "'" Else GrupoPrograma = "null"
        sql = "update MaterialFilmico set GrupoPrograma = " & GrupoPrograma & " where IdMaterial = " & IdMaterial
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function LlenaClientes(ByVal Orden As String, ByVal TipoOrden As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCliente, Cliente, GrupoCliente from Cliente " & _
                "order by " & Orden & " " & TipoOrden

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaGruposCliente() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select '(SIN GRUPO)' as GrupoCliente " & _
                "union select distinct GrupoCliente from Cliente where GrupoCliente is not null " & _
                "order by GrupoCliente"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub ActualizaGrupoCliente(ByVal IdCliente As String, ByVal GrupoCliente As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If GrupoCliente <> "" Then GrupoCliente = "'" & GrupoCliente & "'" Else GrupoCliente = "null"
        sql = "update Cliente set GrupoCliente = " & GrupoCliente & " where IdCliente = " & IdCliente
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function LlenaCentrosCosto(ByVal Orden As String, ByVal TipoOrden As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCentroCosto, CentroCosto, CodTipo, case isnull(CodTipo, 'D') when 'D' then 'Directo' when 'I' then 'Indirecto' else 'Otro' end as Tipo, " & _
                "isnull(FlagDeporte, 'N') as FlagDeporte, CentroCostoEGP, CodGrupoEGP, GrupoEGP " & _
                "from CentroCosto " & _
                "order by " & Orden & " " & TipoOrden

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub ActualizaCentroCosto(ByVal CodCentroCosto As String, ByVal CodTipo As String, ByVal FlagDeporte As Boolean, ByVal CentroCostoEGP As String, ByVal CodGrupoEGP As String, ByVal GrupoEGP As String)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim FlagDeporte2 As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagDeporte Then FlagDeporte2 = "'S'" Else FlagDeporte2 = "null"
        sql = "update CentroCosto set CodTipo = '" & CodTipo & "', FlagDeporte = " & FlagDeporte2 & ", CentroCostoEGP = '" & CentroCostoEGP & "', " & _
                "CodGrupoEGP = " & CodGrupoEGP & ", GrupoEGP = '" & GrupoEGP & "' " & _
                "where CodCentroCosto = " & CodCentroCosto
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Private Function LlenaRating(ByVal IdPeriodo As String) As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtadap As SqlDataAdapter
        Dim dtaset As DataSet

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPeriodo, Programa, Rating, Horas from Rating " & _
                "where IdPeriodo = " & IdPeriodo & " order by Programa"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Sub ProcesaFacturacion(ByVal IdPeriodoCarga As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Procesa_Facturacion"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdPeriodoCarga", IdPeriodoCarga)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Private Sub CreaRating(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        Dim IdAño, IdMes As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "sp_Crea_Rating"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IdAño", IdAño)
        cmd.Parameters.AddWithValue("@IdMes", IdMes)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Private Function BuscaFacturacion(ByVal IdPeriodo As String, ByVal CodGrupo As String, ByVal flagAñoAnt As Boolean) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim IdPeriodoCarga As String
        Dim Facturacion As Double

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Not flagAñoAnt Then IdPeriodoCarga = IdPeriodo Else IdPeriodoCarga = IdPeriodo.Substring(0, 4) & "12"

        sql = "select isnull(sum(MontoUSD), 0) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodoCarga & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = " & CodGrupo

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Facturacion = Convert.ToDouble(rdr("MontoUSD"))
        rdr.Close()
        cn.Close()

        Return Facturacion
    End Function

    Private Function BuscaComisionesAgencias(ByVal IdPeriodo As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim IdAño As String
        Dim Comision As Double

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        IdAño = IdPeriodo.Substring(0, 4)

        sql = "select sum(DebeUSD) - sum(HaberUSD) as MontoTotal " & _
                "from Asiento where Codcuenta in ('953008', '953001102') " & _
                "and year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Comision = Convert.ToDouble(rdr("MontoTotal"))
        rdr.Close()
        cn.Close()

        Return Comision
    End Function

    Public Function BuscaComisionesInternacionales(ByVal IdPeriodo As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim IdAño As String
        Dim Comision As Double

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        IdAño = IdPeriodo.Substring(0, 4)

        sql = "select isnull(sum(MontoUSD), 0) as MontoTotal from V_Movimiento_Ext " & _
            "where CodPersona = 80002 and IdPeriodo between " & IdAño & "01 and " & IdPeriodo

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Comision = Convert.ToDouble(rdr("MontoTotal"))
        rdr.Close()
        cn.Close()

        Return Comision
    End Function

    Sub GeneraReportesVentas(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        'Try

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("VENTAS")
        hoja.Cells(2, 14).Value = IdPeriodo.Substring(4, 2)
        GeneraVentas(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COMP. VENTAS")
        GeneraComparativoVentas(hoja, IdPeriodo)
        hoja = reporte.Worksheets("COMP. VENTAS DET")
        GeneraComparativoVentasDetallado(hoja, IdPeriodo)
        hoja = reporte.Worksheets("SALDO PUBLIC.")
        GeneraSaldoPublicidad(hoja, IdPeriodo)

        For IdMes = 1 To Convert.ToInt32(IdPeriodo.Substring(4, 2))
            GeneraMes(reporte, IdPeriodo, IdMes, "3") 'Cliente
            GeneraMes(reporte, IdPeriodo, IdMes, "1") 'Agencia
            GeneraMes(reporte, IdPeriodo, IdMes, "4") 'Programa
            hoja = reporte.Worksheets("COMISIONES")
            GeneraComisionesAgencias(hoja, IdPeriodo, IdMes)
        Next

        reporte.Worksheets("Plantilla").Delete()
        reporte.Worksheets("FACT. X").Delete()
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

    Private Sub GeneraVentas(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant As Integer

        Dim fil_ini As Integer

        IdAño = IdPeriodo.Substring(0, 4)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CT.IdPeriodo, CT.MontoUSD as Contado, CJ.MontoUSD as Canje, PO.MontoUSD as Politicos, DE.MontoUSD as EventosDeportivos " & _
                "from " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('707001', '7040001', '704000101', '704000102', '704000103', '7591005') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CT, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('707002', '704000201', '704000202') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CJ, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 30 group by IdPeriodo) PO, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 40 group by IdPeriodo) DE " & _
                "where CT.IdPeriodo = CJ.IdPeriodo and CJ.IdPeriodo = PO.IdPeriodo and PO.IdPeriodo = DE.IdPeriodo " & _
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

        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range("C" & fil_ini.ToString & ":C" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 2, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":H" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        hoja.Cells(fil_ini + 13, 4).Value = -1 * BuscaComisionesAgencias(IdPeriodo)
        hoja.Cells(fil_ini + 14, 4).Value = -1 * BuscaComisionesInternacionales(IdPeriodo)

        cn.Close()
    End Sub

    Private Sub GeneraComparativoVentas(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant As Integer

        Dim fil_ini As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CT.IdPeriodo, CT.MontoUSD + CJ.MontoUSD - PO.MontoUSD - DE.MontoUSD as Actual, " & _
                "CT2.MontoUSD + CJ2.MontoUSD - PO2.MontoUSD - DE2.MontoUSD as Anterior " & _
                "from " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('7040001', '704000101', '704000102', '704000103', '7591005') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CT, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta like '7040002%' " & _
                "group by year(Fecha) * 100 + month(Fecha)) CJ, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 30 group by IdPeriodo) PO, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 40 group by IdPeriodo) DE, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " and CodCuenta in  ('7040001', '704000101', '704000102', '704000103', '7591005', '707001', '759004') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CT2, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " and CodCuenta in ('704000201', '704000202', '707002') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CJ2, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdAñoAnt & "12 and IdPeriodo <= " & IdPeriodoAnt & " and CodGrupo = 30 group by IdPeriodo) PO2, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdAñoAnt & "12 and IdPeriodo <= " & IdPeriodoAnt & " and CodGrupo = 40 group by IdPeriodo) DE2 " & _
                "where CT.IdPeriodo = CJ.IdPeriodo and CJ.IdPeriodo = PO.IdPeriodo and PO.IdPeriodo = DE.IdPeriodo and DE.IdPeriodo = CT2.IdPeriodo + 100 " & _
                "and CT2.IdPeriodo = CJ2.IdPeriodo and CJ2.IdPeriodo = PO2.IdPeriodo and PO2.IdPeriodo = DE2.IdPeriodo " & _
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
        hoja.Cells(fil_ini + 6, 3).Value = IdAño
        hoja.Cells(fil_ini + 6, 4).Value = IdAñoAnt

        fil_ini = fil_ini + 7

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        hoja.Cells(fil_ini + 13, 3).Value = BuscaFacturacion(IdPeriodo, "30", False) 'Politicos
        hoja.Cells(fil_ini + 13, 4).Value = BuscaFacturacion(IdPeriodoAnt, "30", True) 'Politicos
        hoja.Cells(fil_ini + 14, 3).Value = BuscaFacturacion(IdPeriodo, "40", False) 'Eventos Deportivos
        hoja.Cells(fil_ini + 14, 4).Value = BuscaFacturacion(IdPeriodoAnt, "40", True) 'Eventos Deportivos

        hoja.Cells(fil_ini + 17, 3).Value = -1 * BuscaComisionesAgencias(IdPeriodo)
        hoja.Cells(fil_ini + 17, 4).Value = -1 * BuscaComisionesAgencias(IdPeriodoAnt)

        hoja.Cells(fil_ini + 18, 3).Value = -1 * BuscaComisionesInternacionales(IdPeriodo)
        hoja.Cells(fil_ini + 18, 4).Value = -1 * BuscaComisionesInternacionales(IdPeriodoAnt)

        cn.Close()
    End Sub

    Private Sub GeneraComparativoVentasDetallado(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant As Integer

        Dim fil_ini As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CT.IdPeriodo, CT.MontoUSD - PO.MontoUSD - DE.MontoUSD as ContadoActual, CT2.MontoUSD - PO2.MontoUSD - DE2.MontoUSD as ContadoAnterior, " & _
                "CJ.MontoUSD as CanjeActual, CJ2.MontoUSD as CanjeAnterior " & _
                "from " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('7040001', '704000101', '704000102', '704000103', '7591005') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CT, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta like '7040002%' " & _
                "group by year(Fecha) * 100 + month(Fecha)) CJ, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 30 group by IdPeriodo) PO, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and IdPeriodo <= " & IdPeriodo & " and CodGrupo = 40 group by IdPeriodo) DE, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " and CodCuenta in  ('7040001', '704000101', '704000102', '704000103', '7591005', '707001', '759004') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CT2, " & _
                "(select year(Fecha) * 100 + month(Fecha) as IdPeriodo, sum(HaberUSD) - sum(DebeUSD) as MontoUSD " & _
                "from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " and CodCuenta in ('704000201', '704000202', '707002') " & _
                "group by year(Fecha) * 100 + month(Fecha)) CJ2, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdAñoAnt & "12 and IdPeriodo <= " & IdPeriodoAnt & " and CodGrupo = 30 group by IdPeriodo) PO2, " & _
                "(select IdPeriodo, sum(MontoUSD) as MontoUSD from Facturacion " & _
                "where IdPeriodoCarga = " & IdAñoAnt & "12 and IdPeriodo <= " & IdPeriodoAnt & " and CodGrupo = 40 group by IdPeriodo) DE2 " & _
                "where CT.IdPeriodo = CJ.IdPeriodo and CJ.IdPeriodo = PO.IdPeriodo and PO.IdPeriodo = DE.IdPeriodo and DE.IdPeriodo = CT2.IdPeriodo + 100 " & _
                "and CT2.IdPeriodo = CJ2.IdPeriodo and CJ2.IdPeriodo = PO2.IdPeriodo and PO2.IdPeriodo = DE2.IdPeriodo " & _
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
        hoja.Cells(fil_ini + 7, 6).Value = IdAño
        hoja.Cells(fil_ini + 7, 7).Value = IdAñoAnt
        hoja.Cells(fil_ini + 7, 9).Value = IdAño
        hoja.Cells(fil_ini + 7, 10).Value = IdAñoAnt

        fil_ini = fil_ini + 8

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":G" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        hoja.Cells(fil_ini + 13, 3).Value = BuscaFacturacion(IdPeriodo, "30", False) 'Politicos
        hoja.Cells(fil_ini + 13, 4).Value = BuscaFacturacion(IdPeriodoAnt, "30", True) 'Politicos
        hoja.Cells(fil_ini + 14, 3).Value = BuscaFacturacion(IdPeriodo, "40", False) 'Eventos Deportivos
        hoja.Cells(fil_ini + 14, 4).Value = BuscaFacturacion(IdPeriodoAnt, "40", True) 'Eventos Deportivos

        hoja.Cells(fil_ini + 17, 3).Value = -1 * BuscaComisionesAgencias(IdPeriodo)
        hoja.Cells(fil_ini + 17, 4).Value = -1 * BuscaComisionesAgencias(IdPeriodoAnt)

        hoja.Cells(fil_ini + 18, 3).Value = -1 * BuscaComisionesInternacionales(IdPeriodo)
        hoja.Cells(fil_ini + 18, 4).Value = -1 * BuscaComisionesInternacionales(IdPeriodoAnt)

        cn.Close()
    End Sub

    Private Sub GeneraSaldoPublicidad(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        'Dim cn As New SqlConnection
        'Dim cmd As New SqlCommand
        'Dim dtaset As DataSet
        'Dim dtadap As SqlDataAdapter
        'Dim dt As DataTable
        'Dim sql As String

        Dim UltDiaMes As Date

        'Dim array As Object(,)
        'Dim cant As Integer

        Dim fil_ini As Integer

        'cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        'cn.Open()

        'sql = ""

        'cmd = New SqlCommand(sql, cn)
        'dtadap = New SqlDataAdapter(cmd)
        'dtaset = New DataSet()
        'dtadap.Fill(dtaset)
        'dt = dtaset.Tables(0)
        'cant = dt.Rows.Count

        fil_ini = 1

        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        fil_ini = fil_ini + 10

        'array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        'hoja.Range("C" & fil_ini.ToString & ":C" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        hoja.Cells(fil_ini + 13, 4).Value = -1 * BuscaComisionesAgencias(IdPeriodo)
        hoja.Cells(fil_ini + 14, 4).Value = -1 * BuscaComisionesInternacionales(IdPeriodo)

        'cn.Close()
    End Sub

    Private Sub GeneraMes(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal IdMes As Integer, ByVal CodGrupo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim IdAño, IdPeriodoMes As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant, i, i2, CantReg As Integer

        Dim fil_ini As Integer

        Dim aux_rating, aux_rating2 As String

        If IdMes = 1 Then
            reporte.Worksheets("FACT. X").Copy(, reporte.Worksheets("SALDO PUBLIC."))
            hoja = reporte.Worksheets("FACT. X (2)")
            If CodGrupo = 1 Then
                hoja.Name = "FACT. X AGENCIA"
                hoja.Columns(16).Hidden = True
            ElseIf CodGrupo = 3 Then
                hoja.Name = "FACT. X CLIENTE"
                hoja.Columns(16).Hidden = True
            Else
                hoja.Name = "FACT. X PROGRAMA"
            End If
        Else
            If CodGrupo = 1 Then
                hoja = reporte.Worksheets("FACT. X AGENCIA")
            ElseIf CodGrupo = 3 Then
                hoja = reporte.Worksheets("FACT. X CLIENTE")
            Else
                hoja = reporte.Worksheets("FACT. X PROGRAMA")
            End If
        End If

        CantReg = 60
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoMes = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If CodGrupo = 4 Then aux_rating = "dbo.f_Rating(" & IdPeriodo & ", " & CodGrupo & ", T.Grupo)" Else aux_rating = "null"
        If CodGrupo = 4 Then aux_rating2 = "case when dbo.f_Rating(" & IdPeriodo & ", " & CodGrupo & ", T.Grupo) is not null then 0 else null end " Else aux_rating2 = "null"
        sql = "select Orden, T.Grupo, isnull(MontoMes, 0) as MontoMes, MontoTotal, " & aux_rating & " as Rating, " & aux_rating2 & " as Rating2 from " & _
                "(select row_number() over (order by sum(MontoUSD) desc) Orden, Grupo, sum(MontoUSD) as MontoTotal " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '*%' " & _
                "and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo " & _
                "union select 1000 as Orden, 'VARIOS' as Grupo, A.MontoTotal - F.MontoTotal as MontoTotal " & _
                "from (select sum(HaberUSD) - sum(DebeUSD) as MontoTotal from Asiento " & _
                "where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('7040001', '704000101', '704000102', '704000103', '7591005', '704000201', '704000202')) A, " & _
                "(select sum(MontoUSD) as MontoTotal from V_Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ") F " & _
                "union select 1000 + row_number() over (order by Grupo), Grupo, sum(MontoUSD) as MontoTotal " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo like '* FACTURADO%' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by Grupo " & _
                ") T left join " & _
                "(select Grupo, sum(MontoUSD) as MontoMes " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' " & _
                "and IdPeriodo = " & IdPeriodoMes & " group by Grupo " & _
                "union select 'VARIOS' as Grupo, A.MontoTotal - F.MontoTotal as MontoTotal " & _
                "from (select sum(HaberUSD) - sum(DebeUSD) as MontoTotal from Asiento A " & _
                "where year(Fecha) * 100 + month(Fecha) = " & IdPeriodoMes & " and CodCuenta in ('7040001', '704000101', '704000102', '704000103', '7591005', '704000201', '704000202')) A, " & _
                "(select sum(MontoUSD) as MontoTotal from V_Facturacion " & _
                "where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' and IdPeriodo = " & IdPeriodoMes & ") F " & _
                ") M on T.Grupo = M.Grupo " & _
                "order by 6 desc, 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        i = 0
        fil_ini = 1

        Do While i < cant
            If IdMes = 1 Then
                rango = reporte.Worksheets("Plantilla").Rows("1:70")
                rango.Copy(hoja.Rows(fil_ini))
                If CodGrupo = 1 Then
                    hoja.Cells(fil_ini + 3, 2).Value = "FACTURACIÓN POR AGENCIAS"
                ElseIf CodGrupo = 3 Then
                    hoja.Cells(fil_ini + 3, 2).Value = "FACTURACIÓN POR CLIENTES"
                Else
                    hoja.Cells(fil_ini + 3, 2).Value = "FACTURACIÓN POR PROGRAMAS"
                    End If
                UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
                hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
                End If

            If i > 0 Then
                hoja.Cells(fil_ini + 7, 2 + IdMes).FormulaR1C1 = "=R[-9]C[0]"
                hoja.Cells(fil_ini + 7, 15).FormulaR1C1 = "=R[-9]C[0]"
            Else
                hoja.Rows(fil_ini + 7).EntireRow.Hidden = True
                End If

            fil_ini = fil_ini + 8

            If i + CantReg < cant Then i2 = i + CantReg - 1 Else i2 = cant - 1

            If IdMes = 1 Then
                array = DataSet2Array(dt, i, i2, 1, 1, -1, -1, -1, -1)
                hoja.Range("B" & fil_ini.ToString & ":B" & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array
                If CodGrupo = 4 Then 'Rating Programas'
                    array = DataSet2Array(dt, i, i2, 1, 4, -1, -1, -1, -1)
                    hoja.Range("P" & fil_ini.ToString & ":P" & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array
                End If
            End If
            array = DataSet2Array(dt, i, i2, 1, 2, -1, -1, -1, -1)
            hoja.Range(Chr(67 + IdMes - 1) & fil_ini.ToString & ":" & Chr(67 + IdMes - 1) & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array
            'array = DataSet2Array(dt, i, i2, 1, 3, -1, -1, -1, -1)
            'hoja.Range("O" & fil_ini.ToString & ":O" & (fil_ini + i2 - i).ToString, Type.Missing).Value2 = array

            If i2 = cant - 1 Then
                hoja.Rows((fil_ini + i2 - i + 1).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True
                hoja.Cells(fil_ini + CantReg, 2).Value = "TOTAL"
                hoja.Cells(6, 15).FormulaR1C1 = "=R" & (fil_ini + CantReg).ToString & "C[0]"
            Else
                hoja.HPageBreaks.Add(hoja.Rows(fil_ini + CantReg + 2))
            End If

            i = i + CantReg
            fil_ini = fil_ini + CantReg + 2
        Loop

        cn.Close()
    End Sub

    Private Sub GeneraComisionesAgencias(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdMes As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdPeriodoMes As String
        Dim UltDiaMes As Date

        Dim array As Object(,)
        Dim cant As Integer

        Dim fil_ini As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoMes = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select T.Persona, isnull(MontoMes, 0) as MontoMes, MontoTotal from " & _
                "(select Persona, sum(DebeUSD) - sum(HaberUSD) as MontoTotal " & _
                "from Asiento where Codcuenta = '953001102' " & _
                "and year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Persona) T left join " & _
                "(select Persona, sum(DebeUSD) - sum(HaberUSD) as MontoMes " & _
                "from Asiento where Codcuenta = '953001102' " & _
                "and year(Fecha) * 100 + month(Fecha) = " & IdPeriodoMes & " " & _
                "group by Persona) M on T.Persona = M.Persona " & _
                "order by 3 desc"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        fil_ini = 1

        If IdMes = 1 Then
            UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        End If

        fil_ini = fil_ini + 7

        If IdMes = 1 Then
            array = DataSet2Array(dt, 1, 0, -1, -1, -1, -1)
            hoja.Range("B" & fil_ini.ToString & ":B" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        End If
        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range(Chr(67 + IdMes - 1) & fil_ini.ToString & ":" & Chr(67 + IdMes - 1) & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array
        'array = DataSet2Array(dt, 1, 2, -1, -1, -1, -1)
        'hoja.Range("O" & fil_ini.ToString & ":O" & (fil_ini + cant - 1).ToString, Type.Missing).Value2 = array

        If cant < 50 Then
            hoja.Rows((fil_ini + cant).ToString & ":" & (fil_ini + 50 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraPlantillaRating(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook, hoja As Excel.Worksheet
        'Try

        FuncionesRentabilidad.DistribucionCostos(0)
        CreaRating(IdPeriodo)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("RATING")

        Dim array As Object(,)
        Dim dt As DataTable
        Dim cant As Integer

        dt = LlenaRating(IdPeriodo)
        cant = dt.Rows.Count

        array = DataSet2Array(dt, False)
        hoja.Range("A2:D" & (cant + 1).ToString, Type.Missing).Value2 = array
        If hoja.Cells(2, 3).Value = "" Then hoja.Cells(2, 3).Value = 0

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

End Module

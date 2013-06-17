Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesPasivos

    Function LlenaPasivos() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdGrupoPasivos, CodSeccion, Seccion, CodGrupo, Grupo, CodSubGrupo, SubGrupo, G.CodCuenta, Cuenta, FlagGpoATV " & _
                "from GrupoPasivos G, CuentaEGP C where G.CodCuenta = C.CodCuenta " & _
                "order by 10, 3, 5, 7, 8"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaSeccion() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodSeccion, Seccion from GrupoPasivos order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaGrupo() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodGrupo, Grupo " & _
                "from GrupoPasivos order by Grupo"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaSubGrupo() As DataTable
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodSubGrupo, SubGrupo " & _
                "from GrupoPasivos order by SubGrupo"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaCuentas_Pasivos() As DataTable
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
                "where CodCuenta like '4%' order by 1"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub CreaPasivo(ByVal CodSeccion As String, ByVal Seccion As String, ByVal CodGrupo As String, ByVal Grupo As String, ByVal CodSubGrupo As String, ByVal SubGrupo As String, ByVal CodCuenta As String, ByVal FlagGpoATV As Boolean)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim FlagGpoATV2 As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagGpoATV Then FlagGpoATV2 = "S" Else FlagGpoATV2 = "N"

        sql = "insert GrupoPasivos(CodSeccion, Seccion, CodGrupo, Grupo, CodSubGrupo, SubGrupo, CodCuenta, FlagGpoATV) " & _
                "values ('" & CodSeccion & "', '" & Seccion & "', " & CodGrupo & ", '" & Grupo & "', " & CodSubGrupo & ", '" & SubGrupo & "', '" & CodCuenta & "', '" & FlagGpoATV2 & "')"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaPasivo(ByVal IdGrupoPasivos As String, ByVal CodSeccion As String, ByVal Seccion As String, ByVal CodGrupo As String, ByVal Grupo As String, ByVal CodSubGrupo As String, ByVal SubGrupo As String, ByVal CodCuenta As String, ByVal FlagGpoATV As Boolean)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand
        Dim FlagGpoATV2 As String

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If FlagGpoATV Then FlagGpoATV2 = "S" Else FlagGpoATV2 = "N"

        sql = "update GrupoPasivos set CodSeccion = '" & CodSeccion & "', Seccion = '" & Seccion & "', CodGrupo = " & CodGrupo & ", Grupo = '" & Grupo & "', " & _
                "CodSubGrupo = " & CodSubGrupo & ", SubGrupo = '" & SubGrupo & "', CodCuenta = '" & CodCuenta & "', FlagGpoATV = '" & FlagGpoATV2 & "' " & _
                "where IdGrupoPasivos = " & IdGrupoPasivos
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaPasivo(ByVal IdGrupoPasivos)
        Dim sql As String
        Dim cn As SqlConnection
        Dim cmd As SqlCommand

        cn = New SqlConnection
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from GrupoPasivos " & _
                "where IdGrupoPasivos = " & IdGrupoPasivos
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function BuscaGrupoPasivo(ByVal CodGrupo As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Grupo As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select top 1 Grupo from GrupoPasivos where CodGrupo = " & CodGrupo
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Grupo = rdr("Grupo").ToString()
        rdr.Close()

        cn.Close()

        Return Grupo
    End Function

    Function BuscaCodSubGrupo(ByVal CodGrupo As String, ByVal SubGrupo As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim CodSubGrupo As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodSubGrupo from GrupoPasivos where CodGrupo = " & CodGrupo & " and SubGrupo = '" & SubGrupo & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then CodSubGrupo = rdr("CodSubGrupo").ToString() Else CodSubGrupo = 0
        rdr.Close()
        cn.Close()

        Return CodSubGrupo
    End Function

    Function BuscaSeccion(ByVal CodSeccion As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Seccion As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select top 1 Seccion from GrupoPasivos where CodSeccion = '" & CodSeccion & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Seccion = rdr("Seccion").ToString()
        rdr.Close()

        cn.Close()

        Return Seccion
    End Function

    Function BuscaGruposPasivos(ByVal IdPeriodo) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        Dim CodSubGrupo, CodSubGrupo2, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        CodSubGrupo = "2"
        CodSubGrupo2 = "3"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)
        IdAñoAnt1 = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt1 = IdAñoAnt1 & "12"

        sql = "select distinct D.CodGrupo, Grupo, CodSeccion from " & _
                "DetallePasivo D, (select distinct CodGrupo, Grupo from GrupoPasivos) G where " & _
                "D.CodGrupo = G.CodGrupo and CodSeccion not in ('PROVED', 'HONORA', 'REMUNE', 'SUNAT') and CodSubGrupo in (2, 3) " & _
                "and (IdPeriodo = " & IdPeriodoAnt & " or (IdPeriodo between " & IdAñoAnt1 & "01 and " & IdPeriodoAnt1 & ") or (IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ")) " & _
                "order by Grupo desc"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaGruposPasivos2() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodSeccion from GrupoPasivos where CodSeccion in ('HONORA', 'REMUNE', 'SUNAT') and CodSubGrupo in (2, 3) order by CodSeccion desc"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaGruposPasivosNoReconocidos(ByVal IdPeriodo) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        Dim CodSubGrupo, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        cn = New SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        CodSubGrupo = "4"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = "1998"
        IdPeriodoAnt1 = "201011"
        IdAñoAnt1 = "2010"

        sql = "select P.CodGrupo, Grupo " & _
                "from " & _
                "(select distinct CodGrupo, Grupo from GrupoPasivos where CodSeccion not in ('HONORA', 'REMUNE') ) P left join " & _
                "(select CodGrupo, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodGrupo) PA1 on P.CodGrupo = PA1.CodGrupo left join " & _
                "(select CodGrupo, sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt1 & "01 and " & IdPeriodoAnt1 & " " & _
                "group by CodGrupo) PA2 on P.CodGrupo = PA2.CodGrupo left join " & _
                "(select CodGrupo, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodGrupo) PA3 on P.CodGrupo = PA3.CodGrupo " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(Monto, 0) <> 0 order by Grupo"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub CreaDetallePasivo()
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_DetallePasivo"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraReportesPasivos(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim UltDiaMes As Date

        'Try
        CreaDetallePasivo()

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        reporte.Worksheets("CARATULA").Cells(45, 5).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper

        GeneraResumenPasivos(reporte, IdPeriodo, 1)
        GeneraResumenIndecopi(reporte, IdPeriodo, 1)
        GeneraResumenCorrientes(reporte, IdPeriodo, 1)
        GeneraResumenNoReconocidos(reporte, IdPeriodo, 1)
        GeneraPasivos(reporte, IdPeriodo)

        reporte.Worksheets("Plantilla").Delete()
        reporte.Worksheets("Plantilla2").Delete()
        reporte.Worksheets("Plantilla3").Delete()
        reporte.Worksheets("Plantilla4").Delete()

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

    Sub GeneraResumenPasivos(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, CodSubGrupo2, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet

        Dim Cant As Integer

        hoja = reporte.Worksheets("RESUMEN")

        CodSubGrupo = "2"
        CodSubGrupo2 = "3"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)
        IdAñoAnt1 = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt1 = IdAñoAnt1 & "12"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 6).Value = "AL " & Convert.ToDateTime(IdAñoAnt1 & "-12-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")
        hoja.Cells(fil_ini + 7, 5).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")
        hoja.Cells(fil_ini + 7, 8).Value = hoja.Cells(fil_ini + 7, 5).Value
        hoja.Cells(fil_ini + 7, 9).Value = hoja.Cells(fil_ini + 7, 5).Value

        fil_ini = fil_ini + 8

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(MontoAnt2, 0) as MontoAnt2, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt2 " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo2 & " and IdPeriodo between " & IdAñoAnt1 & "01 and " & IdPeriodoAnt1 & " " & _
                "group by CodSeccion) PA3 on P.CodSeccion = PA3.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo2 & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA4 on P.CodSeccion = PA4.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(MontoAnt2, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 5, -1, -1, -1, -1)
        hoja.Range("H" & fil_ini.ToString & ":H" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 12 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 12 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 13

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(MontoAnt2, 0) as MontoAnt2, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt2 " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo2 & " and IdPeriodo between " & IdAñoAnt1 & "01 and " & IdPeriodoAnt1 & " " & _
                "group by CodSeccion) PA3 on P.CodSeccion = PA3.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo2 & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA4 on P.CodSeccion = PA4.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(MontoAnt2, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 5, -1, -1, -1, -1)
        hoja.Range("H" & fil_ini.ToString & ":H" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 6 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraResumenIndecopi(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, IdPeriodoAnt, IdAño, IdAñoAnt As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet

        Dim Cant As Integer

        hoja = reporte.Worksheets("RESUMEN PAS.INDECOPI")

        CodSubGrupo = "2"
        IdPeriodoAnt = "199809"
        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 5).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 8

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 12 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 12 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 13

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 6 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraResumenCorrientes(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, IdPeriodoAnt, IdAño, IdAñoAnt As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet

        Dim Cant As Integer

        hoja = reporte.Worksheets("RESUMEN PAS.CORRIENTE")

        CodSubGrupo = "3"
        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & "12"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 3).Value = "AL " & Convert.ToDateTime(IdAñoAnt & "-12-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")
        hoja.Cells(fil_ini + 7, 6).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 8

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 12 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 12 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 13

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 6 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 6 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraResumenNoReconocidos(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet

        Dim Cant As Integer

        hoja = reporte.Worksheets("RESUMEN PAS.NO REC.")

        CodSubGrupo = "4"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = "1998"
        IdPeriodoAnt1 = "201011"
        IdAñoAnt1 = "2010"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 7).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 8

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, -sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo =" & IdPeriodoAnt1 & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'N' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA3 on P.CodSeccion = PA3.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 4, -1, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":G" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 10 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 10 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 11

        sql = "select P.CodSeccion, Seccion, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodSeccion, Seccion from GrupoPasivos) P left join " & _
                "(select CodSeccion, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodSeccion) PA1 on P.CodSeccion = PA1.CodSeccion left join " & _
                "(select CodSeccion, -sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo = " & IdPeriodoAnt1 & " " & _
                "group by CodSeccion) PA2 on P.CodSeccion = PA2.CodSeccion left join " & _
                "(select CodSeccion, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where FlagGpoATV = 'S' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodSeccion) PA3 on P.CodSeccion = PA3.CodSeccion " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(Monto, 0) <> 0 order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("B" & fil_ini.ToString & ":C" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("E" & fil_ini.ToString & ":E" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 4, -1, -1, -1, -1)
        hoja.Range("G" & fil_ini.ToString & ":G" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 10 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 10 - 1).ToString).EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraPasivos(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String)
        Dim hoja As Excel.Worksheet
        Dim IdPeriodoAnt As String
        Dim CodGrupo As Integer
        Dim CodSeccion As String
        Dim fil_ini As Integer
        Dim UltDiaMes As Date
        Dim dt As DataTable

        'Try

        hoja = reporte.Worksheets("PROVEEDORES")
        UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(5, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        fil_ini = 7
        IdPeriodoAnt = "199809"
        GeneraDetalleCorriente(hoja, IdPeriodo, IdPeriodoAnt, "4", "2", 30, fil_ini)
        fil_ini = 63
        IdPeriodoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & "12"
        GeneraDetalleProveedores(reporte, IdPeriodo, IdPeriodoAnt, fil_ini)

        dt = BuscaGruposPasivos2()
        For Each row In dt.Rows
            CodSeccion = row("CodSeccion")
            reporte.Worksheets("Plantilla2").Copy(, reporte.Worksheets("Plantilla2"))
            hoja = reporte.Worksheets("Plantilla2 (2)")
            hoja.Name = BuscaSeccion(CodSeccion)
            hoja.Cells(4, 2).Value = BuscaSeccion(CodSeccion)
            UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(5, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            fil_ini = 7
            IdPeriodoAnt = "199809"
            GeneraDetalleCorriente2(hoja, IdPeriodo, IdPeriodoAnt, CodSeccion, "2", 5, fil_ini)
            fil_ini = 17
            IdPeriodoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & "12"
            GeneraDetalleCorriente2(hoja, IdPeriodo, IdPeriodoAnt, CodSeccion, "3", 5, fil_ini)
        Next

        dt = BuscaGruposPasivos(IdPeriodo)
        For Each row In dt.Rows
            CodGrupo = row("CodGrupo")
            reporte.Worksheets("Plantilla").Copy(, reporte.Worksheets("Plantilla"))
            hoja = reporte.Worksheets("Plantilla (2)")
            hoja.Name = row("Grupo")
            hoja.Cells(4, 2).Value = BuscaGrupoPasivo(CodGrupo)
            UltDiaMes = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(5, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
            fil_ini = 7
            IdPeriodoAnt = "199809"
            GeneraDetalleCorriente(hoja, IdPeriodo, IdPeriodoAnt, CodGrupo, "2", 10, fil_ini)
            fil_ini = 33
            IdPeriodoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & "12"
            GeneraDetalleCorriente(hoja, IdPeriodo, IdPeriodoAnt, CodGrupo, "3", 40, fil_ini)
        Next

        fil_ini = 1
        GeneraDetalleNoReconocidos2(reporte, IdPeriodo, "HONORA", fil_ini)
        GeneraDetalleNoReconocidos2(reporte, IdPeriodo, "REMUNE", fil_ini)
        dt = BuscaGruposPasivosNoReconocidos(IdPeriodo)
        For Each row In dt.Rows
            CodGrupo = row("CodGrupo")
            GeneraDetalleNoReconocidos(reporte, IdPeriodo, CodGrupo, fil_ini)
        Next
        reporte.Worksheets("PAS.NO REC.").PageSetup.PrintArea = "$B$1:$H$" & (fil_ini - 2).ToString

        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'Finally
        'End Try
    End Sub

    Sub GeneraDetalleCorriente(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdPeriodoAnt As String, ByVal CodGrupo As String, ByVal CodSubGrupo As String, ByVal CantReg As Integer, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dr As DataRow()
        Dim sql As String

        Dim IdAño, IdAñoAnt As String

        Dim array As Object(,)

        Dim cant, fil_fin, aux_fil As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)

        If IdPeriodoAnt <> "199809" Then hoja.Cells(fil_ini + 1, 4).Value = "AL 31/12/" & IdAñoAnt Else hoja.Cells(fil_ini + 1, 4).Value = "AL 04/09/1998"
        hoja.Cells(fil_ini + 1, 6).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 3
        aux_fil = fil_ini

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select row_number() over (order by Monto desc, MontoAnt desc) as Orden, P.CodPersona, Persona, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto from " & _
                "(select CodPersona, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodPersona) PA1 full join " & _
                "(select CodPersona, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodPersona) PA2 on PA1.CodPersona = PA2.CodPersona join " & _
                "Persona P on isnull(PA1.CodPersona, PA2.CodPersona) = P.CodPersona " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by 2"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        cant = dtaset.Tables(0).Rows.Count

        If cant > 0 Then
            If cant >= CantReg Then fil_fin = fil_ini + CantReg - 1 Else fil_fin = fil_ini + cant - 1

            dr = dtaset.Tables(0).Select("orden <= " & CantReg.ToString, "Persona")
            array = DataSet2Array(dr, 0, fil_fin - fil_ini, 2, 2, 3, -1, -1, -1)
            hoja.Range("C" & fil_ini.ToString & ":D" & fil_fin.ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dr, 0, fil_fin - fil_ini, 1, 4, -1, -1, -1, -1)
            hoja.Range("F" & fil_ini.ToString & ":F" & fil_fin.ToString, Type.Missing).Value2 = array

            If cant > CantReg Then
                dr = dtaset.Tables(0).Select("orden > " & CantReg.ToString, "Persona")
                fil_ini = fil_fin + 2
                fil_fin = fil_ini + cant - CantReg - 1
                array = DataSet2Array(dr, 0, fil_fin - fil_ini, 2, 2, 3, -1, -1, -1)
                hoja.Range("C" & fil_ini.ToString & ":D" & fil_fin.ToString, Type.Missing).Value2 = array
                array = DataSet2Array(dr, 0, fil_fin - fil_ini, 1, 4, -1, -1, -1, -1)
                hoja.Range("F" & fil_ini.ToString & ":F" & fil_fin.ToString, Type.Missing).Value2 = array
            Else
                hoja.Rows((aux_fil + CantReg).ToString & ":" & (aux_fil + CantReg).ToString).EntireRow.Hidden = True
            End If

            If fil_fin <= aux_fil + CantReg - 1 Then
                hoja.Rows((fil_fin + 1).ToString & ":" & (aux_fil + CantReg - 1).ToString).EntireRow.Hidden = True
            End If
        Else
            If IdPeriodoAnt > "199809" Then hoja.Rows("33:128").EntireRow.Hidden = True Else hoja.Rows("7:32").EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraDetalleCorriente2(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdPeriodoAnt As String, ByVal CodSeccion As String, ByVal CodSubGrupo As String, ByVal CantReg As Integer, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim IdAño, IdAñoAnt As String

        Dim array As Object(,)

        Dim cant, fil_fin As Integer

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)

        If IdPeriodoAnt <> "199809" Then hoja.Cells(fil_ini + 1, 4).Value = "AL 31/12/" & IdAñoAnt Else hoja.Cells(fil_ini + 1, 4).Value = "AL 04/09/1998"
        hoja.Cells(fil_ini + 1, 6).Value = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        fil_ini = fil_ini + 3

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select P.CodGrupo, Grupo, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto from " & _
                "(select CodGrupo, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodSeccion = '" & CodSeccion & "' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodGrupo) PA1 full join " & _
                "(select CodGrupo, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodSeccion = '" & CodSeccion & "' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodGrupo) PA2 on PA1.CodGrupo = PA2.CodGrupo join " & _
                "(select distinct CodGrupo, Grupo from GrupoPasivos) P on isnull(PA1.CodGrupo, PA2.CodGrupo) = P.CodGrupo " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by 2"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        cant = dt.Rows.Count

        If cant > 0 Then
            fil_fin = fil_ini + cant - 1
            array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
            hoja.Range("C" & fil_ini.ToString & ":D" & fil_fin.ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
            hoja.Range("F" & fil_ini.ToString & ":F" & fil_fin.ToString, Type.Missing).Value2 = array
            If fil_fin <= fil_ini + CantReg - 1 Then
                hoja.Rows((fil_fin + 1).ToString & ":" & (fil_ini + CantReg - 1).ToString).EntireRow.Hidden = True
            End If
        Else
            If IdPeriodoAnt > "199809" Then hoja.Rows("17:25").EntireRow.Hidden = True Else hoja.Rows("7:16").EntireRow.Hidden = True
        End If

        cn.Close()
    End Sub

    Sub GeneraDetalleProveedores(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal IdPeriodoAnt As String, ByVal fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dr As DataRow()
        Dim sql As String

        Dim CodGrupo, CodSubGrupo As String
        Dim IdAño, IdAñoAnt As String

        Dim hoja As Excel.Worksheet, rango, rango2 As Excel.Range

        Dim array As Object(,)

        Dim i, cant As Integer

        hoja = reporte.Worksheets("PROVEEDORES")
        CodGrupo = "4"
        CodSubGrupo = "3"
        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = IdPeriodoAnt.Substring(0, 4)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select row_number() over (order by Monto desc, MontoAnt desc) as Orden, P.CodPersona, Persona, isnull(MontoAnt, 0) as MontoAnt, isnull(Monto, 0) as Monto from " & _
                "(select CodPersona, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodPersona) PA1 full join " & _
                "(select CodPersona, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodPersona) PA2 on PA1.CodPersona = PA2.CodPersona join " & _
                "Persona P on isnull(PA1.CodPersona, PA2.CodPersona) = P.CodPersona " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(Monto, 0) <> 0 order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cant = dtaset.Tables(0).Rows.Count
        dr = dtaset.Tables(0).Select("orden <= 200", "Persona")
        i = 0
        Do While i <= 150
            rango = reporte.Worksheets("Plantilla3").Rows("1:59")
            rango.Copy(hoja.Rows(fil_ini))
            hoja.Cells(fil_ini + 5, 2).Value = i
            hoja.Cells(fil_ini + 4, 4).Value = IdAñoAnt & "-12-31"
            hoja.Cells(fil_ini + 4, 6).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")
            If i >= 50 Then
                hoja.Cells(fil_ini + 5, 4).FormulaR1C1 = "=R[-7]C[0]"
                hoja.Cells(fil_ini + 5, 5).FormulaR1C1 = "=R[-7]C[0]"
                hoja.Cells(fil_ini + 5, 6).FormulaR1C1 = "=R[-7]C[0]"
            End If
            fil_ini = fil_ini + 6
            array = DataSet2Array(dr, i, i + 50 - 1, 2, 2, 3, -1, -1, -1)
            hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + 50 - 1).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dr, i, i + 50 - 1, 1, 4, -1, -1, -1, -1)
            hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + 50 - 1).ToString, Type.Missing).Value2 = array
            hoja.HPageBreaks.Add(hoja.Rows(fil_ini + 50 + 2))
            i = i + 50
            fil_ini = fil_ini + 53
        Loop

        If cant > 200 Then
            fil_ini = fil_ini - 2
            dr = dtaset.Tables(0).Select("orden > 200", "Persona")
            rango = reporte.Worksheets("Plantilla3").Rows("7:7")

            rango2 = hoja.Rows(fil_ini.ToString & ":" & (fil_ini + cant - 200 - 1).ToString)
            rango2.Insert()

            rango.Copy(hoja.Rows(fil_ini.ToString & ":" & (fil_ini + cant - 200 - 1).ToString))

            array = DataSet2Array(dr, 2, 2, 3, -1, -1, -1)
            hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array
            array = DataSet2Array(dr, 1, 4, -1, -1, -1, -1)
            hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + dr.GetUpperBound(0)).ToString, Type.Missing).Value2 = array

            rango2 = hoja.Rows(fil_ini.ToString & ":" & (fil_ini + cant - 200 - 1).ToString)
            rango2.EntireRow.Hidden = True

            hoja.Cells(fil_ini - 1, 4).FormulaR1C1 = "=SUM(R[1]C[0]:R[" & (cant - 200).ToString & "]C[0])"
            hoja.Cells(fil_ini - 1, 5).FormulaR1C1 = "=R[0]C[1] - R[0]C[-1]"
            hoja.Cells(fil_ini - 1, 6).FormulaR1C1 = "=SUM(R[1]C[0]:R[" & (cant - 200).ToString & "]C[0])"
        End If

        rango = reporte.Worksheets("Plantilla3").Rows("60")
        rango.Copy(hoja.Rows(fil_ini + cant - 200 + 2))
        hoja.Cells(fil_ini + cant - 200 + 2, 4).FormulaR1C1 = "=R[-2]C[0] + R61C[0]"
        hoja.Cells(fil_ini + cant - 200 + 2, 5).FormulaR1C1 = "=R[-2]C[0] + R61C[0]"
        hoja.Cells(fil_ini + cant - 200 + 2, 6).FormulaR1C1 = "=R[-2]C[0] + R61C[0]"

        'If fil_ini + dtaset.Tables(0).Rows.Count <= fil_ini + 49 Then
        '    hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & (fil_ini + 49).ToString).EntireRow.Hidden = True
        'End If

        cn.Close()
    End Sub

    Sub GeneraDetalleNoReconocidos(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal CodGrupo As String, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim Cant As Integer

        'reporte.Worksheets("Plantilla4").Copy(, reporte.Worksheets("Plantilla4"))
        hoja = reporte.Worksheets("PAS.NO REC.")
        'hoja.Name = "PAS.NO REC."

        CodSubGrupo = "4"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = "1998"
        IdPeriodoAnt1 = "201011"
        IdAñoAnt1 = "2010"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        rango = reporte.Worksheets("Plantilla4").Rows("1:61")
        rango.Copy(hoja.Rows(fil_ini))

        hoja.Cells(fil_ini + 3, 2).Value = BuscaGrupoPasivo(CodGrupo)
        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 8).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 9

        sql = "select P.CodPersona, Persona, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct D.CodPersona, Persona from DetallePasivo D, Persona P where D.CodPersona = P.CodPersona " & _
                "and (IdPeriodo = " & IdPeriodoAnt & " or (IdPeriodo between " & IdAñoAnt1 & "01 and " & IdPeriodoAnt1 & ") or (IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ")) ) P left join " & _
                "(select CodPersona, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodPersona) PA1 on P.CodPersona = PA1.CodPersona left join " & _
                "(select CodPersona, -sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo = " & IdPeriodoAnt1 & " " & _
                "group by CodPersona) PA2 on P.CodPersona = PA2.CodPersona left join " & _
                "(select CodPersona, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodGrupo = " & CodGrupo & " and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodPersona) PA3 on P.CodPersona = PA3.CodPersona " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(Monto, 0) <> 0 order by Persona"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 4, -1, -1, -1, -1)
        hoja.Range("H" & fil_ini.ToString & ":H" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 50 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 50 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 52

        cn.Close()
    End Sub

    Sub GeneraDetalleNoReconocidos2(ByVal reporte As Excel.Workbook, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByRef fil_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim CodSubGrupo, IdPeriodoAnt, IdPeriodoAnt1, IdAño, IdAñoAnt, IdAñoAnt1 As String

        Dim array As Object(,)

        Dim hoja As Excel.Worksheet, rango As Excel.Range

        Dim Cant As Integer

        hoja = reporte.Worksheets("PAS.NO REC.")

        CodSubGrupo = "4"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoAnt = "199809"
        IdAñoAnt = "1998"
        IdPeriodoAnt1 = "201011"
        IdAñoAnt1 = "2010"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        rango = reporte.Worksheets("Plantilla4").Rows("1:61")
        rango.Copy(hoja.Rows(fil_ini))
        hoja.Cells(fil_ini + 3, 2).Value = BuscaSeccion(CodSeccion)
        Dim UltDiaMes As Date = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
        hoja.Cells(fil_ini + 4, 2).Value = "AL " & UltDiaMes.ToString("dd 'de' MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE")).ToUpper
        hoja.Cells(fil_ini + 7, 8).Value = "AL " & Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("dd/MM/yyyy")

        fil_ini = fil_ini + 9

        sql = "select P.CodGrupo, Grupo, isnull(MontoAnt, 0) as MontoAnt, isnull(MontoAnt1, 0) as MontoAnt1, isnull(Monto, 0) as Monto " & _
                "from " & _
                "(select distinct CodGrupo, Grupo from GrupoPasivos) P left join " & _
                "(select CodGrupo, sum(MontoUSD) as MontoAnt " & _
                "from DetallePasivo where CodSeccion = '" & CodSeccion & "' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " " & _
                "group by CodGrupo) PA1 on P.CodGrupo = PA1.CodGrupo left join " & _
                "(select CodGrupo, -sum(MontoUSD) as MontoAnt1 " & _
                "from DetallePasivo where CodSeccion = '" & CodSeccion & "' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo = " & IdPeriodoAnt1 & " " & _
                "group by CodGrupo) PA2 on P.CodGrupo = PA2.CodGrupo left join " & _
                "(select CodGrupo, sum(MontoUSD) as Monto " & _
                "from DetallePasivo where CodSeccion = '" & CodSeccion & "' and CodSubGrupo = " & CodSubGrupo & " and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by CodGrupo) PA3 on P.CodGrupo = PA3.CodGrupo " & _
                "where isnull(MontoAnt, 0) <> 0 or isnull(MontoAnt1, 0) <> 0 or isnull(Monto, 0) <> 0 order by Grupo"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)
        Cant = dt.Rows.Count

        array = DataSet2Array(dt, 2, 1, 2, -1, -1, -1)
        hoja.Range("C" & fil_ini.ToString & ":D" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 3, -1, -1, -1, -1)
        hoja.Range("F" & fil_ini.ToString & ":F" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array
        array = DataSet2Array(dt, 1, 4, -1, -1, -1, -1)
        hoja.Range("H" & fil_ini.ToString & ":H" & (fil_ini + Cant - 1).ToString, Type.Missing).Value2 = array

        If fil_ini + Cant <= fil_ini + 50 - 1 Then
            hoja.Rows((fil_ini + Cant).ToString & ":" & (fil_ini + 50 - 1).ToString).EntireRow.Hidden = True
        End If

        fil_ini = fil_ini + 52

        cn.Close()
    End Sub

End Module

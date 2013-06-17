Imports Microsoft.Office.Interop
Imports System.Configuration
Imports System.Data.SqlClient

Module FuncionesRep

    Function LlenaAños() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct substring(convert(varchar(6), IdPeriodo), 1, 4) as IdAño from Periodo order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaAños2011() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct substring(convert(varchar(6), IdPeriodo), 1, 4) as IdAño " & _
                "from Periodo where IdPeriodo >= 201101 order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaPeriodos() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPeriodo, Periodo from Periodo order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaPeriodosVentas() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPeriodo, Periodo from Periodo where IdPeriodo >= 201112 order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaPeriodos2011() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdPeriodo, Periodo from Periodo where IdPeriodo >= 201101 order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function LlenaCtasBancos() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()
        sql = "select 0 as CodCuentaBanco, '-- TODAS --' as CuentaBanco " & _
                "union select CodCuentaBanco, convert(varchar(2), CodCuentaBanco) + ' - ' + Cuenta as CuentaBanco " & _
                "from CuentaContable where CodCuentaBanco is not null order by 1"
        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cn.Close()

        Return dtaset.Tables(0)
    End Function

    Function BuscaPeriodo(ByVal IdPeriodo As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Periodo As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()
        sql = "select Periodo from Periodo where IdPeriodo = " & IdPeriodo
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Periodo = rdr("Periodo").ToString()
        rdr.Close()
        cn.Close()

        Return Periodo
    End Function

    Function BuscaMontoNeto(ByVal IdPeriodo As String, ByVal CodCuentaFC2 As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim MontoNeto As Double = 0

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select sum(MontoBaseUSD) as MontoNeto from V_Movimiento " & _
                "where IdPeriodo = " & IdPeriodo & " and CodCuentaFC2 = " + CodCuentaFC2
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() And Not IsDBNull(rdr("MontoNeto")) Then MontoNeto = Convert.ToDouble(rdr("MontoNeto"))
        rdr.Close()

        cn.Close()

        Return MontoNeto
    End Function

    Function BuscaSaldoFinalContab(ByVal IdPeriodo As String) As Double
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim SaldoFinal As Double

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select sum(Saldo) as Saldo " & _
                "from V_SaldoBancos where IdAño * 100 + IdMes = " & IdPeriodo & " and NroCuenta like '10%' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() And Not IsDBNull(rdr("Saldo")) Then SaldoFinal = Convert.ToDouble(rdr("Saldo"))
        rdr.Close()

        cn.Close()

        Return SaldoFinal
    End Function

    Function BuscaMes(ByVal IdMes As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rdr As SqlDataReader
        Dim Mes As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()
        sql = "select Mes from Mes where IdMes = " & IdMes
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Mes = rdr("Mes").ToString().ToUpper()
        rdr.Close()
        cn.Close()

        Return Mes
    End Function

    'Function BuscaITFX(ByVal IdPeriodoIni As String, ByVal IdPeriodoFin As String) As Double
    '    Dim sql As String
    '    Dim cn As New OleDbConnection
    '    Dim cmd As New OleDbCommand
    '    Dim rdr As OleDbDataReader
    '    Dim ITF As Double = 0

    '    cn = New OleDbConnection()
    '    cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
    '    cn.Open()

    '    sql = "select sum(isnull(DebeUSD, 0) - isnull(HaberUSD, 0)) as MontoUSD " & _
    '            "from LibroMayor where NroCuenta = '954000104' and IdAño * 100 + IdMes between " & IdPeriodoIni & " and " & IdPeriodoFin
    '    cmd = New OleDbCommand(sql, cn)
    '    rdr = cmd.ExecuteReader
    '    If rdr.Read() And Not IsDBNull(rdr("MontoUSD")) Then ITF = Convert.ToDouble(rdr("MontoUSD"))
    '    rdr.Close()

    '    cn.Close()

    '    Return ITF
    'End Function

    Sub LlenaITF()
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from Ajuste where CodCuentaFC2 = 902"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "insert into Ajuste(IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) " & _
                "select IdAño * 100 + IdMes as IdPeriodo, convert(smalldatetime, convert(varchar(10), IdAño * 10000 + IdMes * 100 + 1)) as Fecha, " & _
                "902 as CodCuentaFC2, sum(isnull(DebeUSD, 0) - isnull(HaberUSD, 0)) as MontoUSD " & _
                "from LibroMayor where NroCuenta = '954000104' " & _
                "group by IdAño * 100 + IdMes, convert(smalldatetime, convert(varchar(10), IdAño * 10000 + IdMes * 100 + 1))"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "update Ajuste set MontoUSD = MontoUSD - ITFTesoreria " & _
                "from Ajuste A, (select IdPeriodo, sum(MontoUSD) as ITFTesoreria from Movimiento " & _
                "where CodCuentaFC = 96 group by IdPeriodo) M " & _
                "where A.IdPeriodo = M.IdPeriodo and A.CodCuentaFC2 = 902"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ExportarExcel(ByVal Response As System.Web.HttpResponse, ByVal Grid As System.Web.UI.WebControls.GridView, ByVal Archivo As String)
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        Grid.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(Grid)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=" & Archivo)
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.Default
        Response.Write(sw.ToString())
        Response.End()
    End Sub

End Module

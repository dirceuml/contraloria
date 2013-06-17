Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesInfGestion

    Function LlenaER() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaER, CtaER, isnull(Signo, 0) as Signo, isnull(CodSeccion, '') as CodSeccion, isnull(FlagModif, 'N') as FlagModif, Orden " & _
                "from CtaER " & _
                "order by Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaERExp() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Orden, CtaER, CodSeccion, CodCuenta, Cuenta, " & _
                "case D.Signo when 1 then '+' when -1 then '-' else '' end as Signo " & _
                "from CtaER C left join CtaERDet D on C.IdCtaER = D.IdCtaER left join CuentaEGP CC on D.CodCtaOrigen = CC.CodCuenta " & _
                "order by Orden, CtaER, CodCuenta"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaCtaER(ByVal IdCtaER As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim CtaER As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CtaER from CtaER where IdCtaER = " & IdCtaER
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        CtaER = rdr("CtaER").ToString()
        rdr.Close()

        cn.Close()

        Return CtaER
    End Function

    Function LlenaCodSeccionER() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select '' as CodSeccion union select distinct CodSeccion from CtaER where FlagModif = 'S' " & _
                "order by CodSeccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaERDet(ByVal IdCtaER As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaERDet, Signo, CodCtaOrigen, CodCuenta + ' ' + Cuenta as Cuenta " & _
                "from CtaERDet D, CuentaEGP C " & _
                "where D.CodCtaOrigen = C.CodCuenta and IdCtaER = " & IdCtaER & " " & _
                "order by 3"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If dt.Rows.Count = 0 Then
            sql = "select 0 as IdCtaERDet, 1 as Signo, null as CodCtaOrigen, null as Cuenta"
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)
        End If

        cn.Close()

        Return dt
    End Function

    Function LlenaCuentasER() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCuenta, CodCuenta + ' ' + Cuenta as Cuenta " & _
                "from CuentaEGP " & _
                "where convert(tinyint, substring(CodCuenta, 1, 2)) in (67, 70, 75, 76, 77, 95, 99)"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub AgregaCtaER(ByVal CtaER As String, ByVal Signo As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Signo = "0" Then Signo = "null"
        If CodSeccion <> "" Then CodSeccion = "'" & CodSeccion & "'" Else CodSeccion = "null"
        If FlagModif = "S" Then FlagModif = "'S'" Else FlagModif = "null"

        sql = "insert into CtaER(CtaER, Signo, CodSeccion, FlagModif, Orden) " & _
                "values ('" & CtaER & "', " & Signo & ", " & CodSeccion & ", " & FlagModif & ", " & Orden & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaER(ByVal IdCtaER As String, ByVal CtaER As String, ByVal Signo As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If Signo = "0" Then Signo = "null"
        If CodSeccion <> "" Then CodSeccion = "'" & CodSeccion & "'" Else CodSeccion = "null"
        If FlagModif = "S" Then FlagModif = "'S'" Else FlagModif = "null"

        sql = "update CtaER set CtaER = '" & CtaER & "', Signo = " & Signo & ", CodSeccion = " & CodSeccion & ", FlagModif = " & FlagModif & ", Orden = " & Orden & " " & _
                "where IdCtaER =" & IdCtaER
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaER(ByVal IdCtaER As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaERDet where IdCtaER = " & IdCtaER
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "delete from CtaER where IdCtaER = " & IdCtaER
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub AgregaCtaERDet(ByVal IdCtaER As String, ByVal Signo As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into CtaERDet(IdCtaER, CodCtaOrigen, Signo) " & _
                "values (" & IdCtaER & ", " & CodCtaOrigen & ", " & Signo & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaERDet(ByVal IdCtaERDet As String, ByVal Signo As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaERDet set CodCtaOrigen = " & CodCtaOrigen & ", Signo = " & Signo & " " & _
                "where IdCtaERDet =" & IdCtaERDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaERDet(ByVal IdCtaERDet As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaERDet where IdCtaERDet = " & IdCtaERDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Private Sub CreaDetalleER(ByVal IdPeriodo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        Dim IdAño, IdMes As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "sp_Crea_DetalleER"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IdAño", IdAño)
        cmd.Parameters.AddWithValue("@IdMes", IdMes)
        cmd.ExecuteNonQuery()

        sql = "sp_Crea_DetalleERPpto"
        cmd = New SqlCommand(sql, cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IdAño", IdAño)
        cmd.Parameters.AddWithValue("@IdMes", IdMes)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function LlenaBG() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaBG, CtaBG, CodSeccion, isnull(FlagModif, 'N') as FlagModif, Orden from CtaBG " & _
                "order by Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaBGExp() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Orden, CtaBG, CodSeccion, CodCuenta, Cuenta " & _
                "from CtaBG C left join CtaBGDet D on C.IdCtaBG = D.IdCtaBG left join CuentaEGP CC on D.CodCtaOrigen = CC.CodCuenta " & _
                "order by Orden, CtaBG, CodCuenta"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaCtaBG(ByVal IdCtaBG As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim CtaBG As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CtaBG from CtaBG where IdCtaBG = " & IdCtaBG
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        CtaBG = rdr("CtaBG").ToString()
        rdr.Close()

        cn.Close()

        Return CtaBG
    End Function

    Function LlenaCodSeccionBG() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodSeccion from CtaBG where FlagModif = 'S' " & _
                "order by CodSeccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaBGDet(ByVal IdCtaBG As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaBGDet, CodCuenta + ' ' + Cuenta as Cuenta " & _
                "from CtaBGDet D, CuentaEGP C " & _
                "where D.CodCtaOrigen = C.CodCuenta and IdCtaBG = " & IdCtaBG & " " & _
                "order by 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaCuentasBG() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCuenta, CodCuenta + ' ' + Cuenta as Cuenta " & _
                "from CuentaEGP " & _
                "where convert(tinyint, substring(CodCuenta, 1, 2)) between 10 and 59 order by 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub AgregaCtaBG(ByVal CtaBG As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into CtaBG(CtaBG, CodSeccion, FlagModif, Orden) " & _
                "values ('" & CtaBG & "', '" & CodSeccion & "', '" & FlagModif & "', " & Orden & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaBG(ByVal IdCtaBG As String, ByVal CtaBG As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaBG set CtaBG = '" & CtaBG & "', CodSeccion = '" & CodSeccion & "', FlagModif = '" & FlagModif & "', Orden = " & Orden & " " & _
                "where IdCtaBG =" & IdCtaBG
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaBG(ByVal IdCtaBG As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaBGDet where IdCtaBG = " & IdCtaBG
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "delete from CtaBG where IdCtaBG = " & IdCtaBG
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub AgregaCtaBGDet(ByVal IdCtaBGDet As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into CtaBGDet(IdCtaBGDet, CodCtaOrigen) " & _
                "values (" & IdCtaBGDet & ", " & CodCtaOrigen & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaBGDet(ByVal IdCtaBGDet As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaBGDet set CodCtaOrigen = " & CodCtaOrigen & " " & _
                "where IdCtaBGDet =" & IdCtaBGDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaBGDet(ByVal IdCtaBGDet As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaBGDet where IdCtaBGDet = " & IdCtaBGDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Private Sub CreaDetalleBG(ByVal IdPeriodo As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Dim IdPeriodoEne As Integer

        IdPeriodoEne = Convert.ToInt32(IdPeriodo.ToString.Substring(0, 4) & "01")

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_DetalleBG"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdPeriodoEne", IdPeriodoEne)
            cmd.Parameters.AddWithValue("@IdPeriodo", IdPeriodo)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Function LlenaEOAF() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaEOAF, CtaEOAF, Signo, CodSeccion, isnull(FlagModif, 'N') as FlagModif, Orden from CtaEOAF " & _
                "order by Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaEOAFExp() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Orden, CtaEOAF, CodSeccion, Cuenta2 + ' [' + convert(varchar, CodCuenta2) + ']' as CuentaFC, " & _
                "case D.Signo when 1 then '+' when -1 then '-' else '' end as Signo " & _
                "from CtaEOAF C left join CtaEOAFDet D on C.IdCtaEOAF = D.IdCtaEOAF left join V_CuentaFC F on D.CodCtaOrigen = F.CodCuenta2 " & _
                "order by Orden, CtaEOAF, CuentaFC "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaCtaEOAF(ByVal IdCtaEOAF As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim CtaEOAF As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CtaEOAF from CtaEOAF where IdCtaEOAF = " & IdCtaEOAF
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        CtaEOAF = rdr("CtaEOAF").ToString()
        rdr.Close()

        cn.Close()

        Return CtaEOAF
    End Function

    Function LlenaCodSeccionEOAF() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct CodSeccion from CtaEOAF where FlagModif = 'S' " & _
                "order by CodSeccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function LlenaEOAFDet(ByVal IdCtaOEAF As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdCtaEOAFDet, Signo, CodCtaOrigen, Cuenta2 + ' [' + convert(varchar, CodCuenta2) + ']' as CuentaFC " & _
                "from CtaEOAFDet D, V_CuentaFC C " & _
                "where D.CodCtaOrigen = C.CodCuenta2 and IdCtaEOAF = " & IdCtaOEAF & " " & _
                "order by 4, 3"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If dt.Rows.Count = 0 Then
            sql = "select 0 as IdCtaEOAFDet, 1 as Signo, null as CodCtaOrigen, null as CuentaFC"
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)
        End If

        cn.Close()

        Return dt
    End Function

    Function LlenaCuentasFC() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CodCuenta2, Cuenta2 + ' [' + convert(varchar, CodCuenta2) + ']' as Cuenta2 " & _
                "from V_CuentaFC order by 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Sub AgregaCtaEOAF(ByVal CtaEOAF As String, ByVal Signo As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into CtaEOAF(CtaEOAF, Signo, CodSeccion, FlagModif, Orden) " & _
                "values ('" & CtaEOAF & "', " & Signo & ", '" & CodSeccion & "', '" & FlagModif & "', " & Orden & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaEOAF(ByVal IdCtaEOAF As String, ByVal CtaEOAF As String, ByVal Signo As String, ByVal CodSeccion As String, ByVal FlagModif As String, ByVal Orden As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaEOAF set CtaEOAF = '" & CtaEOAF & "', Signo = " & Signo & ", CodSeccion = '" & CodSeccion & "', FlagModif = '" & FlagModif & "', Orden = " & Orden & " " & _
                "where IdCtaEOAF =" & IdCtaEOAF
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaEOAF(ByVal IdCtaEOAF As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaEOAFDet where IdCtaEOAF = " & IdCtaEOAF
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        sql = "delete from CtaEOAF where IdCtaEOAF = " & IdCtaEOAF
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub AgregaCtaEOAFDet(ByVal IdCtaEOAF As String, ByVal Signo As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into CtaEOAFDet(IdCtaEOAF, CodCtaOrigen, Signo) " & _
                "values (" & IdCtaEOAF & ", " & CodCtaOrigen & ", " & Signo & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaCtaEOAFDet(ByVal IdCtaEOAFDet As String, ByVal Signo As String, ByVal CodCtaOrigen As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update CtaEOAFDet set CodCtaOrigen = " & CodCtaOrigen & ", Signo = " & Signo & " " & _
                "where IdCtaEOAFDet =" & IdCtaEOAFDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaCtaEOAFDet(ByVal IdCtaEOAFDet As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from CtaEOAFDet where IdCtaEOAFDet = " & IdCtaEOAFDet
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Private Sub CreaDetalleEOAF(ByVal IdPeriodo As Integer)
        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim sql As String

        Dim IdPeriodoEne As Integer

        IdPeriodoEne = Convert.ToInt32(IdPeriodo.ToString.Substring(0, 4) & "01")

        Try
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
            cn.Open()

            sql = "sp_Crea_DetalleEOAF"
            cmd = New SqlCommand(sql, cn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdPeriodoEne", IdPeriodoEne)
            cmd.Parameters.AddWithValue("@IdPeriodo", IdPeriodo)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            cn.Close()
        End Try

    End Sub

    Sub GeneraReportesInfGestion(ByVal RutaPlantilla As String, ByVal RutaArchivo As String, ByVal IdPeriodo As String)
        Dim excel As Excel.Application
        Dim reporte As Excel.Workbook
        Dim hoja As Excel.Worksheet

        'Try
        FuncionesEGP.CreaDetalleEGP(IdPeriodo)
        FuncionesEvalEGP.CreaDetalleEGPPpto(IdPeriodo)
        CreaDetalleER(Convert.ToInt32(IdPeriodo))
        CreaDetalleEOAF(Convert.ToInt32(IdPeriodo))
        CreaDetalleBG(Convert.ToInt32(IdPeriodo))
        FuncionesRentabilidad.DistribucionCostos(1)

        excel = New Excel.Application
        excel.DisplayAlerts = False
        reporte = excel.Workbooks.Open(RutaPlantilla)

        hoja = reporte.Worksheets("A")
        GeneraER_A(hoja, IdPeriodo)

        hoja = reporte.Worksheets("B")
        GeneraER_B(hoja, IdPeriodo)

        hoja = reporte.Worksheets("C")
        GeneraER_C(hoja, IdPeriodo)

        hoja = reporte.Worksheets("D")
        GeneraER_D(hoja, IdPeriodo)

        hoja = reporte.Worksheets("E")
        GeneraEOAF(hoja, IdPeriodo)

        hoja = reporte.Worksheets("F")
        GeneraBG(hoja, IdPeriodo)

        hoja = reporte.Worksheets("G")
        GeneraVentas(hoja, IdPeriodo)

        hoja = reporte.Worksheets("H")
        GeneraCostosDirectos(hoja, IdPeriodo)

        hoja = reporte.Worksheets("I")
        GeneraContribucion(hoja, IdPeriodo, 11, 11 + 70 - 1, 3)

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

    Private Sub GeneraER_A(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño As Integer
        Dim fil_ini, col_ini As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))

        hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        hoja.Cells(8, 4).Value = BuscaPeriodo(IdPeriodo).Replace(IdAño.ToString, "")

        hoja.Cells(9, 4).Value = IdAño
        hoja.Cells(9, 5).Value = IdAño - 1
        hoja.Cells(9, 9).Value = IdAño
        hoja.Cells(9, 10).Value = IdAño - 1

        col_ini = 3
        fil_ini = 11
        GeneraRubrosER(hoja, IdPeriodo, "VENTAS", fil_ini, fil_ini + 5 - 1, col_ini)
        fil_ini = 19
        GeneraCentrosCostoER(hoja, IdPeriodo, "PRG", fil_ini, fil_ini + 4 - 1, col_ini)
        fil_ini = 26
        GeneraCentrosCostoER(hoja, IdPeriodo, "NOT", fil_ini, fil_ini + 5 - 1, col_ini)
        fil_ini = 34
        GeneraCentrosCostoER(hoja, IdPeriodo, "NOV", fil_ini, fil_ini + 2 - 1, col_ini)
        fil_ini = 39
        GeneraCentrosCostoER(hoja, IdPeriodo, "IND", fil_ini, fil_ini + 20 - 1, col_ini)
        fil_ini = 65
        GeneraRubrosER(hoja, IdPeriodo, "CANJES", fil_ini, fil_ini + 10 - 1, col_ini)
        fil_ini = 79
        GeneraRubrosER(hoja, IdPeriodo, "OTRRUB", fil_ini, fil_ini + 10 - 1, col_ini)
    End Sub

    Private Sub GeneraER_B(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño As Integer
        Dim fil_ini, col_ini As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))

        hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        hoja.Cells(8, 4).Value = BuscaPeriodo(IdPeriodo).Replace(IdAño.ToString, "")

        hoja.Cells(9, 4).Value = IdAño
        hoja.Cells(9, 5).Value = IdAño - 1
        hoja.Cells(9, 9).Value = IdAño
        hoja.Cells(9, 10).Value = IdAño - 1

        col_ini = 3
        fil_ini = 11
        GeneraRubrosER(hoja, IdPeriodo, "VENTAS", fil_ini, fil_ini + 5 - 1, col_ini)
    End Sub

    Private Sub GeneraER_C(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño, IdMes As Integer
        Dim IdPeriodoAnt As String
        Dim fil_ini, col_ini As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
        IdPeriodoAnt = (IdAño - 1).ToString & IdPeriodo.Substring(4, 2)

        If IdMes > 1 Then
            hoja.Cells(4, 3).Value = "Enero a " & BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        Else
            hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        End If

        hoja.Cells(8, 4).Value = IdAño
        hoja.Cells(8, 17).Value = IdAño
        hoja.Cells(8, 19).Value = IdAño - 1

        col_ini = 3
        For i = 1 To IdMes
            fil_ini = 11
            GeneraRubrosER(hoja, IdPeriodo, "VENTAS", i, fil_ini, fil_ini + 5 - 1, col_ini, col_ini + i)
            GeneraRubrosER(hoja, IdPeriodoAnt, "VENTAS", IdMes, fil_ini, fil_ini + 5 - 1, col_ini, col_ini + 16, True)
            fil_ini = 19
            GeneraCentrosCostoER(hoja, IdPeriodo, "PRG", i, fil_ini, fil_ini + 4 - 1, col_ini, col_ini + i)
            GeneraCentrosCostoER(hoja, IdPeriodoAnt, "PRG", i, fil_ini, fil_ini + 4 - 1, col_ini, col_ini + 16, True)
            fil_ini = 26
            GeneraCentrosCostoER(hoja, IdPeriodo, "NOT", i, fil_ini, fil_ini + 5 - 1, col_ini, col_ini + i)
            GeneraCentrosCostoER(hoja, IdPeriodoAnt, "NOT", i, fil_ini, fil_ini + 5 - 1, col_ini, col_ini + 16, True)
            fil_ini = 34
            GeneraCentrosCostoER(hoja, IdPeriodo, "NOV", i, fil_ini, fil_ini + 2 - 1, col_ini, col_ini + i)
            GeneraCentrosCostoER(hoja, IdPeriodoAnt, "NOV", i, fil_ini, fil_ini + 2 - 1, col_ini, col_ini + 16, True)
            fil_ini = 39
            GeneraCentrosCostoER(hoja, IdPeriodo, "IND", i, fil_ini, fil_ini + 20 - 1, col_ini, col_ini + i)
            GeneraCentrosCostoER(hoja, IdPeriodoAnt, "IND", i, fil_ini, fil_ini + 20 - 1, col_ini, col_ini + 16, True)
            fil_ini = 65
            GeneraRubrosER(hoja, IdPeriodo, "CANJES", i, fil_ini, fil_ini + 10 - 1, col_ini, col_ini + i)
            GeneraRubrosER(hoja, IdPeriodoAnt, "CANJES", i, fil_ini, fil_ini + 10 - 1, col_ini, col_ini + 16, True)
            fil_ini = 79
            GeneraRubrosER(hoja, IdPeriodo, "OTRRUB", i, fil_ini, fil_ini + 10 - 1, col_ini, col_ini + i)
            GeneraRubrosER(hoja, IdPeriodoAnt, "OTRRUB", i, fil_ini, fil_ini + 10 - 1, col_ini, col_ini + 16, True)
        Next

        If IdMes < 12 Then
            hoja.Columns(Chr(65 + 3 + IdMes) & ":" & Chr(65 + 3 + 12 - 1)).EntireColumn.Hidden = True
        End If
    End Sub

    Private Sub GeneraER_D(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño As Integer
        Dim fil_ini, col_ini As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))

        hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        hoja.Cells(8, 4).Value = BuscaPeriodo(IdPeriodo).Replace(IdAño.ToString, "")

        col_ini = 3
        fil_ini = 11
        GeneraRubrosERPpto(hoja, IdPeriodo, "VENTAS", fil_ini, fil_ini + 5 - 1, col_ini)
        fil_ini = 19
        GeneraCentrosCostoERPpto(hoja, IdPeriodo, "PRG", fil_ini, fil_ini + 4 - 1, col_ini)
        fil_ini = 26
        GeneraCentrosCostoERPpto(hoja, IdPeriodo, "NOT", fil_ini, fil_ini + 5 - 1, col_ini)
        fil_ini = 34
        GeneraCentrosCostoERPpto(hoja, IdPeriodo, "NOV", fil_ini, fil_ini + 2 - 1, col_ini)
        fil_ini = 39
        GeneraCentrosCostoERPpto(hoja, IdPeriodo, "IND", fil_ini, fil_ini + 20 - 1, col_ini)
        fil_ini = 65
        GeneraRubrosERPpto(hoja, IdPeriodo, "CANJES", fil_ini, fil_ini + 10 - 1, col_ini)
        fil_ini = 79
        GeneraRubrosERPpto(hoja, IdPeriodo, "OTRRUB", fil_ini, fil_ini + 10 - 1, col_ini)
    End Sub

    Private Sub GeneraRubrosER(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.CtaER, isnull(D.Monto, 0) as Monto, isnull(D1.MontoAnt, 0) as MontoAnt, isnull(D2.Acum, 0) as Acum, isnull(D3.AcumAnt, 0) as AcumAnt " & _
                "from CtaER C left join " & _
                "(select IdCtaER, sum(MontoUSD) as Monto from DetalleER " & _
                "where IdPeriodo = " & IdPeriodo & " group by IdCtaER) D on C.IdCtaER = D.IdCtaER left join " & _
                "(select IdCtaER, sum(MontoUSD) as MontoAnt from DetalleER " & _
                "where IdPeriodo = " & IdPeriodoAnt & " group by IdCtaER) D1 on C.IdCtaER = D1.IdCtaER left join" & _
                "(select IdCtaER, sum(MontoUSD) as Acum from DetalleER " & _
                "where IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaER) D2 on C.IdCtaER = D2.IdCtaER left join " & _
                "(select IdCtaER, sum(MontoUSD) as AcumAnt from DetalleER " & _
                "where IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " group by IdCtaER) D3 on C.IdCtaER = D3.IdCtaER " & _
                "where C.CodSeccion = '" & CodSeccion & "' order by C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 2)).Value2 = array

        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + 6), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 7)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraCentrosCostoER(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipoER As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CC.CentroCostoER, isnull(D.Monto, 0) as Monto, isnull(D1.MontoAnt, 0) as MontoAnt, isnull(D2.Acum, 0) as Acum, isnull(D3.AcumAnt, 0) as AcumAnt " & _
                "from V_CentroCostoER CC left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as Monto from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo = " & IdPeriodo & " group by CodCentroCosto) D on CC.CodCentroCosto = D.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoAnt from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo = " & IdPeriodoAnt & " group by CodCentroCosto) D1 on CC.CodCentroCosto = D1.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as Acum from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) D2 on CC.CodCentroCosto = D2.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as AcumAnt from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo between " & IdAñoAnt & "01 and " & IdPeriodoAnt & " group by CodCentroCosto) D3 on CC.CodCentroCosto = D3.CodCentroCosto " & _
                "where CC.CodTipoER = '" & CodTipoER & "' "
        If CodTipoER = "PRG" Then sql = sql & "order by CC.CodCentroCosto" Else sql = sql & "order by CC.CentroCostoER "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 2)).Value2 = array

        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + 6), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 7)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraRubrosER(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal IdMes As Integer, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, ByVal col_mes As Integer, Optional ByVal flagAcum As Boolean = False)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdPeriodoIni, IdPeriodoFin As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoFin = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString
        If Not flagAcum Then IdPeriodoIni = IdPeriodoFin Else IdPeriodoIni = IdAño & "01"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.CtaER, isnull(MontoUSD, 0) as MontoUSD " & _
                "from CtaER C left join " & _
                "(select IdCtaER, sum(MontoUSD) as MontoUSD from DetalleER " & _
                "where IdPeriodo between " & IdPeriodoIni & " and " & IdPeriodoFin & " group by IdCtaER) D " & _
                "on C.IdCtaER = D.IdCtaER where C.CodSeccion = '" & CodSeccion & "' order by C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If IdMes = 1 Then
            array = DataSet2Array(dt, 1, 0, -1, -1, -1, -1)
            hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini)).Value2 = array
        End If

        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_mes), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_mes)).Value2 = array

        If IdMes = 1 And fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraCentrosCostoER(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipoER As String, ByVal IdMes As Integer, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, ByVal col_mes As Integer, Optional ByVal flagAcum As Boolean = False)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdPeriodoIni, IdPeriodoFin As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoFin = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString
        If Not flagAcum Then IdPeriodoIni = IdPeriodoFin Else IdPeriodoIni = IdAño & "01"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CC.CentroCostoER, isnull(MontoUSD, 0) as MontoUSD " & _
                "from V_CentroCostoER CC left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as MontoUSD from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo between " & IdPeriodoIni & " and " & IdPeriodoFin & " group by CodCentroCosto) D " & _
                "on CC.CodCentroCosto = D.CodCentroCosto where CC.CodTipoER = '" & CodTipoER & "' "
        If CodTipoER = "PRG" Then sql = sql & "order by CC.CodCentroCosto" Else sql = sql & "order by CC.CentroCostoER "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If IdMes = 1 Then
            array = DataSet2Array(dt, 1, 0, -1, -1, -1, -1)
            hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini)).Value2 = array
        End If

        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_mes), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_mes)).Value2 = array

        If IdMes = 1 And fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraRubrosERPpto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.CtaER, isnull(D.Monto, 0) as Monto, isnull(D.Ppto, 0) as Ppto, isnull(D1.Acum, 0) as Acum, isnull(D1.AcumPpto, 0) as AcumPpto " & _
                "from CtaER C left join " & _
                "(select IdCtaER, sum(MontoUSD) as Monto, sum(PptoUSD) as Ppto from DetalleER " & _
                "where IdPeriodo = " & IdPeriodo & " group by IdCtaER) D on C.IdCtaER = D.IdCtaER left join " & _
                "(select IdCtaER, sum(MontoUSD) as Acum, sum(PptoUSD) as AcumPpto from DetalleER " & _
                "where IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by IdCtaER) D1 on C.IdCtaER = D1.IdCtaER " & _
                "where C.CodSeccion = '" & CodSeccion & "' order by C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 2)).Value2 = array

        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + 6), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 7)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraCentrosCostoERPpto(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodTipoER As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdAñoAnt, IdPeriodoAnt As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdAñoAnt = (Convert.ToInt32(IdAño) - 1).ToString
        IdPeriodoAnt = IdAñoAnt & IdPeriodo.Substring(4, 2)

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select CC.CentroCostoER, isnull(D.Monto, 0) as Monto, isnull(D.Ppto, 0) as Ppto, isnull(D1.Acum, 0) as Acum, isnull(D1.AcumPpto, 0) as AcumPpto " & _
                "from V_CentroCostoER CC left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as Monto, sum(PptoUSD) as Ppto from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo = " & IdPeriodo & " group by CodCentroCosto) D on CC.CodCentroCosto = D.CodCentroCosto left join " & _
                "(select CodCentroCosto, sum(MontoUSD) as Acum, sum(PptoUSD) as AcumPpto from DetalleER D, CtaER E " & _
                "where D.IdCtaER = E.IdCtaER and E.CodSeccion = 'COSTOS' " & _
                "and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by CodCentroCosto) D1 on CC.CodCentroCosto = D1.CodCentroCosto " & _
                "where CC.CodTipoER = '" & CodTipoER & "' "
        If CodTipoER = "PRG" Then sql = sql & "order by CC.CodCentroCosto" Else sql = sql & "order by CC.CentroCostoER "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        array = DataSet2Array(dt, 3, 0, 1, 2, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 2)).Value2 = array

        array = DataSet2Array(dt, 2, 3, 4, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + 6), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 7)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraEOAF(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño, IdMes As Integer
        Dim IdAñoMes As Integer
        Dim UltDiaMes As Date

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
        IdAñoMes = IdAño * 100

        If IdMes > 1 Then
            hoja.Cells(4, 3).Value = "Enero a " & BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        Else
            hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        End If

        hoja.Cells(47, 4).Value = BuscaCierreCaja((IdAño - 1).ToString & "-12-31")
        For i = 1 To IdMes
            IdAñoMes = IdAñoMes + 1
            GeneraRubrosEOAF(hoja, IdPeriodo, "COBRAN", IdAñoMes.ToString, 10, 14, 3)
            GeneraRubrosEOAF(hoja, IdPeriodo, "REMUNE", IdAñoMes.ToString, 17, 17, 3)
            GeneraRubrosEOAF(hoja, IdPeriodo, "SEGSOC", IdAñoMes.ToString, 19, 23, 3)
            GeneraRubrosEOAF(hoja, IdPeriodo, "IMPTOS", IdAñoMes.ToString, 25, 29, 3)
            GeneraRubrosEOAF(hoja, IdPeriodo, "PROVED", IdAñoMes.ToString, 31, 35, 3)
            GeneraRubrosEOAF(hoja, IdPeriodo, "OTRPAG", IdAñoMes.ToString, 39, 43, 3)
            UltDiaMes = Convert.ToDateTime(IdAñoMes.ToString.Substring(0, 4) & "-" & IdAñoMes.ToString.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1)
            hoja.Cells(52, 3 + i).Value = BuscaCierreCaja(UltDiaMes)
        Next
        hoja.Cells(47, 3 + IdMes + 1).Value = 0
        hoja.Cells(52, 17).FormulaR1C1 = "=R[0]C" & (3 + IdMes).ToString
        If IdMes < 12 Then
            hoja.Columns(Chr(65 + 3 + IdMes) & ":" & Chr(65 + 3 + 11)).EntireColumn.Hidden = True
        End If
    End Sub

    Private Sub GeneraRubrosEOAF(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMesAux As Integer

        Dim array As Object(,)

        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.CtaEOAF, case when CodSeccion = 'COBRAN' then Signo * isnull(MontoBruto, 0) else -1 * Signo * isnull(MontoBruto, 0) end as MontoBruto " & _
                "from CtaEOAF C left join (select * from DetalleEOAF where IdPeriodo = " & IdPeriodoAux & ") D " & _
                "on C.IdCtaEOAF = D.IdCtaEOAF where C.CodSeccion = '" & CodSeccion & "' order by C.Orden"

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

    Private Sub GeneraBG(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdAño, IdMes As Integer
        Dim IdAñoMes As Integer
        Dim fil_ini, col_ini As Integer

        IdAño = Convert.ToInt32(IdPeriodo.Substring(0, 4))
        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))
        IdAñoMes = IdAño * 100

        If IdMes > 1 Then
            hoja.Cells(4, 3).Value = "Enero a " & BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        Else
            hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        End If

        col_ini = 3
        For i = 1 To IdMes
            IdAñoMes = IdAñoMes + 1
            fil_ini = 10
            GeneraRubrosBG(hoja, IdPeriodo, "ACTCOR", IdAñoMes.ToString, fil_ini, fil_ini + 10 - 1, col_ini)
            fil_ini = 21
            GeneraRubrosBG(hoja, IdPeriodo, "ACTNOC", IdAñoMes.ToString, fil_ini, fil_ini + 5 - 1, col_ini)
            fil_ini = 31
            GeneraRubrosBG(hoja, IdPeriodo, "PASCOR", IdAñoMes.ToString, fil_ini, fil_ini + 5 - 1, col_ini)
            fil_ini = 37
            GeneraRubrosBG(hoja, IdPeriodo, "PASNOC", IdAñoMes.ToString, fil_ini, fil_ini + 10 - 1, col_ini)
            fil_ini = 51
            GeneraRubrosBG(hoja, IdPeriodo, "PATRIM", IdAñoMes.ToString, fil_ini, fil_ini + 1 - 1, col_ini)
        Next
        If IdMes < 12 Then
            hoja.Columns(Chr(65 + col_ini + IdMes + 1) & ":" & Chr(65 + col_ini + 12)).EntireColumn.Hidden = True
        End If
    End Sub

    Private Sub GeneraRubrosBG(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal CodSeccion As String, ByVal IdPeriodoAux As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim sql As String

        Dim IdMesAux As Integer

        Dim array As Object(,)

        IdMesAux = Convert.ToInt32(IdPeriodoAux.Substring(4, 2))

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select C.CtaBG, MontoUSD " & _
                "from CtaBG C left join (select * from DetalleBG D where IdPeriodo = " & IdPeriodoAux & ") D " & _
                "on C.IdCtaBG = D.IdCtaBG where C.CodSeccion = '" & CodSeccion & "' order by C.Orden"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        array = DataSet2Array(dtaset.Tables(0), 1, 0, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini)).Value2 = array
        array = DataSet2Array(dtaset.Tables(0), 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMesAux + 1), hoja.Cells(fil_ini + dtaset.Tables(0).Rows.Count - 1, col_ini + IdMesAux + 1)).Value2 = array

        If fil_fin > fil_ini + dtaset.Tables(0).Rows.Count - 1 Then hoja.Rows((fil_ini + dtaset.Tables(0).Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraVentas(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdMes As Integer
        Dim fil_ini, col_ini As Integer

        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        If IdMes > 1 Then
            hoja.Cells(4, 3).Value = "Enero a " & BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        Else
            hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        End If

        col_ini = 3
        For i = 1 To IdMes
            fil_ini = 10
            GeneraVentasMes(hoja, IdPeriodo, i, fil_ini, fil_ini + 60 - 1, col_ini)
        Next
        If IdMes < 12 Then
            hoja.Columns(Chr(65 - 1 + col_ini + IdMes + 1) & ":" & Chr(65 - 1 + col_ini + 12)).EntireColumn.Hidden = True
        End If
    End Sub

    Private Sub GeneraVentasMes(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdMes As Integer, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        Dim array As Object(,)

        Dim CodGrupo As String
        Dim IdAño, IdPeriodoMes As String

        CodGrupo = "4"
        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoMes = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Orden, T.Grupo, isnull(MontoMes, 0) as MontoMes, MontoAcum from " & _
                "(select 1 as Orden, Grupo, sum(MontoUSD) as MontoAcum " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '*%' " & _
                "and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " " & _
                "group by Grupo "

        '"union select 100 as Orden, 'VARIOS' as Grupo, A.MontoAcum - F.MontoAcum as MontoAcum "
        '"from (select sum(HaberUSD) - sum(DebeUSD) as MontoAcum from Asiento " & _
        '"where year(Fecha) * 100 + month(Fecha) between " & IdAño & "01 and " & IdPeriodo & " and CodCuenta in ('7040001', '704000201', '704000202')) A, " & _
        '"(select sum(MontoUSD) as MontoAcum from V_Facturacion " & _
        '"where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & ") F " & _
        '"union select 200 as Orden, Grupo, sum(MontoUSD) as MontoAcum " & _
        '"from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo like '* FACTURADO%' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by Grupo " & _

        sql = sql & ") T left join " & _
                "(select Grupo, sum(MontoUSD) as MontoMes " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' " & _
                "and IdPeriodo = " & IdPeriodoMes & " group by Grupo "

        '"union select 'VARIOS' as Grupo, A.MontoAcum - F.MontoAcum as MontoAcum " & _
        '"from (select sum(HaberUSD) - sum(DebeUSD) as MontoAcum from Asiento A " & _
        '"where year(Fecha) * 100 + month(Fecha) = " & IdPeriodoMes & " and CodCuenta in ('7040001', '704000201', '704000202')) A, " & _
        '"(select sum(MontoUSD) as MontoAcum from V_Facturacion " & _
        '"where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' and IdPeriodo = " & IdPeriodoMes & ") F " & _

        sql = sql & ") M on T.Grupo = M.Grupo " & _
                "order by 2" '1, 4 desc, 2"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If IdMes = 1 Then
            array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
            hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini)).Value2 = array
        End If
        array = DataSet2Array(dt, 1, 2, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_ini + IdMes), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + IdMes)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Private Sub GeneraCostosDirectos(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String)
        Dim IdMes As Integer
        Dim fil_ini, col_ini As Integer

        IdMes = Convert.ToInt32(IdPeriodo.Substring(4, 2))

        If IdMes > 1 Then
            hoja.Cells(4, 3).Value = "Enero a " & BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        Else
            hoja.Cells(4, 3).Value = BuscaPeriodo(IdPeriodo) & " - Cifras en US$"
        End If

        col_ini = 3
        For i = 1 To IdMes
            fil_ini = 10
            GeneraCostosDirectosMes(hoja, IdPeriodo, i, fil_ini, fil_ini + 60 - 1, col_ini, col_ini + i)
        Next
        If IdMes < 12 Then
            hoja.Columns(Chr(65 - 1 + col_ini + IdMes + 1) & ":" & Chr(65 - 1 + col_ini + 12)).EntireColumn.Hidden = True
        End If
    End Sub

    Private Sub GeneraCostosDirectosMes(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal IdMes As Integer, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer, ByVal col_mes As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim IdAño, IdPeriodoMes As String

        IdAño = IdPeriodo.Substring(0, 4)
        IdPeriodoMes = (Convert.ToInt32(IdAño) * 100 + IdMes).ToString

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select upper(D.GrupoPrograma) as GrupoPrograma, isnull(Monto, 0) as Monto, isnull(Acum, 0) as Acum " & _
                "from (select GrupoPrograma, sum(CostoDirecto) as Acum from Distribucion " & _
                "where GrupoPrograma <> 'Material Filmico' and IdPeriodo between " & IdAño & "01 and " & IdPeriodo & " group by GrupoPrograma) D left join " & _
                "(select GrupoPrograma, sum(CostoDirecto) as Monto from Distribucion " & _
                "where GrupoPrograma <> 'Material Filmico' and IdPeriodo = " & IdPeriodoMes & " group by GrupoPrograma) D1 on D.GrupoPrograma = D1.GrupoPrograma " & _
                "order by 1 "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If IdMes = 1 Then
            array = DataSet2Array(dt, 1, 0, -1, -1, -1, -1)
            hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini)).Value2 = array
        End If

        array = DataSet2Array(dt, 1, 1, -1, -1, -1, -1)
        hoja.Range(hoja.Cells(fil_ini, col_mes), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_mes)).Value2 = array

        If IdMes = 1 And fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        cn.Close()
    End Sub

    Function BuscaCostosIndirectos(ByVal IdPeriodo As String, ByVal CodTipoER As String, ByVal CodCentroCosto As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select -1 * isnull(sum(MontoUSD), 0) as Monto from DetalleER D, CtaER E, V_CentroCostoER CC " & _
                "where D.IdCtaER = E.IdCtaER and D.CodCentroCosto = CC.CodCentroCosto " & _
                "and E.CodSeccion = 'COSTOS' and CC.CodTipoER = '" & CodTipoER & "'" & _
                "and IdPeriodo = " & IdPeriodo & " "
        If CodCentroCosto <> "" Then sql = sql & "and D.CodCentroCosto = " & CodCentroCosto

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        Return Convert.ToDouble(dt.Rows(0)("Monto"))
        cn.Close()
    End Function

    Private Sub GeneraContribucion(ByVal hoja As Excel.Worksheet, ByVal IdPeriodo As String, ByVal fil_ini As Integer, ByVal fil_fin As Integer, ByVal col_ini As Integer)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable
        Dim sql As String
        Dim array As Object(,)

        Dim CodGrupo As String

        CodGrupo = "4"

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Rating, Horas, upper(isnull(Grupo, GrupoPrograma)) as GrupoPrograma, isnull(Venta, 0) as Venta, isnull(CostoDirecto, 0) as CostoDirecto " & _
                "from (select Grupo, sum(MontoUSD) as Venta " & _
                "from V_Facturacion where IdPeriodoCarga = " & IdPeriodo & " and CodGrupo = " & CodGrupo & " and Grupo <> 'SIN CONSUMO' and Grupo <> 'SIN PROGRAMA' and Grupo not like '* SIN CONTRATO%' and Grupo not like '* NC ANULAN%' " & _
                "and IdPeriodo = " & IdPeriodo & " group by Grupo) V full join " & _
                "(select GrupoPrograma, sum(CostoDirecto) as CostoDirecto from Distribucion " & _
                "where GrupoPrograma not in ('Material Filmico', 'Pagina Web') and IdPeriodo = " & IdPeriodo & " group by GrupoPrograma) C on V.Grupo = C.GrupoPrograma left join " & _
                "(select Programa, Rating, Horas from Rating where IdPeriodo = " & IdPeriodo & ") R on isnull(V.Grupo, C.GrupoPrograma) = R.Programa " & _
                "order by 3 "

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        array = DataSet2Array(dt, False)
        hoja.Range(hoja.Cells(fil_ini, col_ini), hoja.Cells(fil_ini + dt.Rows.Count - 1, col_ini + 4)).Value2 = array

        If fil_fin > fil_ini + dt.Rows.Count - 1 Then hoja.Rows((fil_ini + dt.Rows.Count).ToString & ":" & fil_fin.ToString).EntireRow.Hidden = True

        hoja.Cells(84, 8).Value = BuscaCostosIndirectos(IdPeriodo, "PRG", 201) + BuscaCostosIndirectos(IdPeriodo, "PRG", 203)
        hoja.Cells(84, 12).Value = BuscaCostosIndirectos(IdPeriodo, "IND", 1) + BuscaCostosIndirectos(IdPeriodo, "IND", 3)
        hoja.Cells(84, 14).Value = BuscaCostosIndirectos(IdPeriodo, "IND", "") - hoja.Cells(84, 12).Value
        cn.Close()
    End Sub

End Module

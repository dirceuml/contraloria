Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Module FuncionesSeguridad

    Function ValidaUsuario(ByVal CodUsuario As String, ByVal Password As String, ByRef IdUsuario As String) As Boolean
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim flag As Boolean

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select count(*) as Cant from Usuario where CodUsuario = '" & CodUsuario & "' and Password = '" & Password & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        flag = (Convert.ToInt32(rdr("Cant") = 1))
        rdr.Close()
        If flag Then
            sql = "select IdUsuario from Usuario where CodUsuario = '" & CodUsuario & "' and Password = '" & Password & "' "
            cmd = New SqlCommand(sql, cn)
            rdr = cmd.ExecuteReader
            rdr.Read()
            IdUsuario = rdr("IdUsuario").ToString
            rdr.Close()
        End If

        cn.Close()

        Return flag
    End Function

    Function ValidaAcceso(ByVal IdUsuario As String, ByVal Pagina As String) As Boolean
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim flag As Boolean

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select count(*) as Cant from Seguridad S, Acceso A " & _
                "where S.IdAcceso = A.IdAcceso and IdUsuario = " & IdUsuario & " and Pagina = '" & Pagina & "' "
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        flag = (Convert.ToInt32(rdr("Cant") = 1))
        rdr.Close()

        cn.Close()

        Return flag
    End Function

    Function LlenaAccesos() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdAcceso, Seccion, Orden, Acceso, Seccion + ' - ' + Acceso as SeccionAcceso, Pagina " & _
                "from Acceso " & _
                "order by Seccion, Orden, Acceso"

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
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select distinct Seccion " & _
                "from Acceso order by Seccion"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function GeneraSubMenu(ByVal IdUsuario As String, ByVal Seccion As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim Acceso, Pagina As String
        Dim SubMenu As String = ""

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        If IdUsuario = "" Then 'ADM
            sql = "select Acceso, Pagina from Acceso " & _
                    "where Seccion = '" & Seccion & "' order by Orden"
        Else
            sql = "select Acceso, Pagina from Acceso A, Seguridad S " & _
                    "where A.IdAcceso = S.IdAcceso and A.Seccion = '" & Seccion & "' " & _
                    "and S.IdUsuario = " & IdUsuario & " order by Orden"
        End If

        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        While rdr.Read()
            Acceso = rdr("Acceso").ToString()
            Pagina = rdr("Pagina").ToString()
            SubMenu = SubMenu & "<li><a href='" & Pagina & "'>" & Acceso & "</a></li>"
        End While
        rdr.Close()

        cn.Close()

        Return SubMenu
    End Function

    Public Function GenerarMenu(ByVal Username As String, ByVal Seccion As String) As String
        Dim menu As String = ""
        Dim submenu As String

        If Seccion = "Importación" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"

        ElseIf Seccion = "Actualización" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"
        ElseIf Seccion = "Mantenimientos" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"
        ElseIf Seccion = "Reportes" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"
        ElseIf Seccion = "Consultas" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"
        ElseIf Seccion = "Seguridad" Then
            menu = "<li><a href='#'>" & Seccion & " +</a><ul>"
            submenu = GeneraSubMenu(Username, Seccion)
            menu = menu & submenu & "</ul><li>"
        End If

        Return menu
    End Function

    Sub AgregaAcceso(ByVal Seccion As String, ByVal Orden As String, ByVal Acceso As String, ByVal Pagina As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into Acceso(Seccion, Orden, Acceso, Pagina) " & _
                "values ('" & Seccion & "', " & Orden & ", '" & Acceso & "', '" & Pagina & "')"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaAcceso(ByVal IdAcceso As String, ByVal Seccion As String, ByVal Orden As String, ByVal Acceso As String, ByVal Pagina As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update Acceso set Seccion = '" & Seccion & "', Orden = " & Orden & ", Acceso = '" & Acceso & "', Pagina = '" & Pagina & "' " & _
                "where IdAcceso = " & IdAcceso
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaAcceso(ByVal IdAcceso As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from Acceso where IdAcceso = " & IdAcceso
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function LlenaUsuarios() As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select * from Usuario order by Usuario"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        cn.Close()

        Return dt
    End Function

    Function BuscaUsuario(ByVal IdUsuario As String, ByRef CodTipo As String) As String
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim Usuario As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select Usuario, CodTipo from Usuario where IdUsuario = " & IdUsuario
        cmd = New SqlCommand(sql, cn)
        rdr = cmd.ExecuteReader
        rdr.Read()
        Usuario = rdr("Usuario").ToString()
        CodTipo = rdr("CodTipo").ToString()
        rdr.Close()

        cn.Close()

        Return Usuario
    End Function

    Sub AgregaUsuario(ByVal Usuario As String, ByVal CodUsuario As String, ByVal Password As String, ByVal CodTipo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into Usuario(Usuario, CodUsuario, Password, CodTipo) " & _
                "values ('" & Usuario & "', '" & CodUsuario & "', '" & Password & "', '" & CodTipo & "')"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub ActualizaUsuario(ByVal IdUsuario As String, ByVal Usuario As String, ByVal CodUsuario As String, ByVal Password As String, ByVal CodTipo As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "update Usuario set Usuario = '" & Usuario & "', CodUsuario = '" & CodUsuario & "', Password = '" & Password & "', CodTipo = '" & CodTipo & "' " & _
                "where IdUsuario = " & IdUsuario
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaUsuario(ByVal IdUsuario As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from Usuario where IdUsuario = " & IdUsuario
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Function LlenaSeguridad(ByVal IdUsuario As String) As DataTable
        Dim sql As String
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim dtaset As DataSet
        Dim dtadap As SqlDataAdapter
        Dim dt As DataTable

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "select IdSeguridad, Seccion, Seccion + ' - ' + Acceso as Acceso " & _
                "from Seguridad S, Acceso A " & _
                "where S.IdAcceso = A.IdAcceso and IdUsuario = " & IdUsuario & " order by Seccion, Orden, Acceso"

        cmd = New SqlCommand(sql, cn)
        dtadap = New SqlDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        dt = dtaset.Tables(0)

        If dt.Rows.Count = 0 Then
            sql = "select 0 as IdSeguridad, null as Acceso"
            cmd = New SqlCommand(sql, cn)
            dtadap = New SqlDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            dt = dtaset.Tables(0)
        End If

        cn.Close()

        Return dt
    End Function

    Sub AgregaSeguridad(ByVal IdUsuario As String, ByVal IdAcceso As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into Seguridad(IdUsuario, IdAcceso) " & _
                "values (" & IdUsuario & ", " & IdAcceso & ")"
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub EliminaSeguridad(ByVal IdSeguridad As String)
        Dim cn As New SqlConnection
        Dim cmd As SqlCommand
        Dim sql As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "delete from Seguridad where IdSeguridad = " & IdSeguridad
        cmd = New SqlCommand(sql, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

End Module

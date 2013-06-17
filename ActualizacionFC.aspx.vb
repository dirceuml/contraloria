Imports System.Data.OleDb
Imports System.Configuration

Public Class ActualizacionFC
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ActualizacionFC.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargarCombos()
            cboPeriodo.SelectedValue = Year(Date.Now) * 100 + Month(Date.Now)
            hdfPagina.Value = 0
            gdvMovimiento.PageSize = Funciones.iREGxPAG
            LlenaGridMovimiento(0)
        End If
        If hdfAccion.Value = "Recargar" Then
            LlenaGridMovimiento(0)
            hdfAccion.Value = ""
        End If
    End Sub

    Sub CargarCombos()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = ""
        sql += "select 0 as IdPeriodo, '-- TODO --' as Periodo "
        sql += "UNION "
        sql += "SELECT IdPeriodo, Periodo FROM Periodo WHERE IdPeriodo >= 201001 "
        sql += "ORDER BY 1 "

        sql += "select 0 as CodCuenta, ' -- TODAS LAS CUENTAS -- ' AS Cuenta "
        sql += "UNION SELECT CodCuenta, '[' + convert(varchar(10), CodCuenta) + '] - ' + Cuenta as Cuenta FROM CuentaBanco "
        sql += "ORDER BY 1 "

        sql += "select 0 as CodCuenta, '-- TODAS LAS CUENTAS --' AS Cuenta "
        sql += "UNION SELECT CodCuenta, Cuenta + ' - [' + convert(varchar(10), CodCuenta) + ']' as Cuenta FROM CuentaFC "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodCuenta2, '-- TODAS LAS CUENTAS --' AS Cuenta2 "
        sql += "UNION "
        sql += "select -2 as CodCuenta2, '--ASIGNADAS--' AS Cuenta2 "
        sql += "UNION "
        sql += "select -1 as CodCuenta2, '--SIN ASIGNAR--' AS Cuenta2 "
        sql += "UNION SELECT CodCuenta2, Cuenta2 + ' - [' + convert(varchar(10), CodCuenta2) + ']' as Cuenta2 FROM CuentaFC2 "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodPersona, '-- TODAS LAS PERSONAS --' AS Persona "
        sql += "UNION SELECT CodPersona, Persona FROM Persona "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodArea, '-- TODAS LAS AREAS --' AS Area "
        sql += "UNION SELECT CodArea, Area + ' - [' + convert(varchar(10), CodArea) + ']' as Area FROM Area "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodArea, '-- TODAS LAS AREAS --' AS Area "
        sql += "UNION "
        sql += "select -2 as CodArea, '--ASIGNADAS--' AS Area "
        sql += "UNION "
        sql += "select -1 as CodArea, '--SIN ASIGNAR--' AS Area "
        sql += "UNION SELECT CodArea, Area + ' - [' + convert(varchar(10), CodArea) + ']' as Area FROM Area "
        sql += "ORDER BY 2 "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboPeriodo.DataSource = dtaset.Tables(0)
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataBind()

        cboCuentaBanco.DataSource = dtaset.Tables(1)
        cboCuentaBanco.DataValueField = "CodCuenta"
        cboCuentaBanco.DataTextField = "Cuenta"
        cboCuentaBanco.DataBind()

        cboCuentaFC.DataSource = dtaset.Tables(2)
        cboCuentaFC.DataValueField = "CodCuenta"
        cboCuentaFC.DataTextField = "Cuenta"
        cboCuentaFC.DataBind()

        cboCuentaFCNueva.DataSource = dtaset.Tables(3)
        cboCuentaFCNueva.DataValueField = "CodCuenta2"
        cboCuentaFCNueva.DataTextField = "Cuenta2"
        cboCuentaFCNueva.DataBind()

        cboPersona.DataSource = dtaset.Tables(4)
        cboPersona.DataValueField = "CodPersona"
        cboPersona.DataTextField = "Persona"
        cboPersona.DataBind()

        cboArea.DataSource = dtaset.Tables(5)
        cboArea.DataValueField = "CodArea"
        cboArea.DataTextField = "Area"
        cboArea.DataBind()

        cboAreaNueva.DataSource = dtaset.Tables(6)
        cboAreaNueva.DataValueField = "CodArea"
        cboAreaNueva.DataTextField = "Area"
        cboAreaNueva.DataBind()

        cn.Close()

    End Sub

    Sub LlenaGridMovimiento(ByVal pagina As Integer)
        Dim sSql1, sSql2 As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        'Unimos los filtros
        sSql2 = "From Movimiento WHERE IdMovimiento > 0 "
        If cboPeriodo.Text <> "0" Then
            sSql2 += " AND IdPeriodo = " + cboPeriodo.Text + " "
        End If
        If cboCuentaBanco.Text <> "0" Then
            sSql2 += " AND CodCuentaBanco = " + cboCuentaBanco.Text + " "
        End If
        If cboCuentaFC.Text <> "0" Then
            sSql2 += " AND CodCuentaFC = " + cboCuentaFC.Text + " "
        End If
        If cboCuentaFCNueva.Text = "-1" Then
            sSql2 += " AND IdMovimiento NOT IN (SELECT IdMovimiento FROM Movimiento2 M2, Movimiento M WHERE M2.IdPeriodo = M.IdPeriodo AND M2.CodCuentaBanco = M.CodCuentaBanco AND M2.NroVoucher = M.NroVoucher AND M2.NroItem = M.NroItem AND M2.CodCuentaFC = M.CodCuentaFC AND M2.Observaciones = M.Observaciones) "
        ElseIf cboCuentaFCNueva.Text = "-2" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM Movimiento2 M2, Movimiento M WHERE M2.IdPeriodo = M.IdPeriodo AND M2.CodCuentaBanco = M.CodCuentaBanco AND M2.NroVoucher = M.NroVoucher AND M2.NroItem = M.NroItem AND M2.CodCuentaFC = M.CodCuentaFC AND M2.Observaciones = M.Observaciones) "
        ElseIf cboCuentaFCNueva.Text <> "0" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM Movimiento2 M2, Movimiento M WHERE M2.IdPeriodo = M.IdPeriodo AND M2.CodCuentaBanco = M.CodCuentaBanco AND M2.NroVoucher = M.NroVoucher AND M2.NroItem = M.NroItem AND M2.CodCuentaFC = M.CodCuentaFC AND M2.Observaciones = M.Observaciones AND M2.CodCuentaFC2 = " + cboCuentaFCNueva.Text + ") "
        End If
        If txtPersona.Text <> "" Then
            sSql2 += " AND Persona like '%" + txtPersona.Text + "%' "
        End If
        If cboPersona.Text <> "0" Then
            sSql2 += " AND CodPersona = " + cboPersona.Text + " "
        End If
        If cboArea.Text = "-1" Then
            sSql2 += " AND CodArea IS NULL "
        ElseIf cboArea.Text <> "0" Then
            sSql2 += " AND CodArea = " + cboArea.Text + " "
        End If
        If cboAreaNueva.Text <> "0" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM Movimiento2 M2, Movimiento M WHERE M2.IdPeriodo = M.IdPeriodo AND M2.CodCuentaBanco = M.CodCuentaBanco AND M2.NroVoucher = M.NroVoucher AND M2.NroItem = M.NroItem AND M2.CodCuentaFC = M.CodCuentaFC AND M2.Observaciones = M.Observaciones AND M2.CodArea2 = " + cboAreaNueva.Text + ") "
        End If
        If txtFecha.Text <> "" Then
            sSql2 += " AND Fecha = '" + txtFecha.Text + "' "
        End If
        If txtNroVoucher.Text <> "" Then
            sSql2 += " AND NroVoucher = " + txtNroVoucher.Text + " "
        End If
        If txtGlosa.Text <> "" Then
            sSql2 += " AND Glosa like '%" + txtGlosa.Text + "%' "
        End If

        'Extraemos el numero de registros
        hdfNroRegistros.Value = Funciones.ExtraerValor("Select Count(*) as NumReg " + sSql2, "NumReg")
        If hdfNroRegistros.Value <> "0" Then
            lblNumRegistros.Text = "Se han encontrado " + hdfNroRegistros.Value + " registros"
            'LLenamos la grila
            sSql1 = " Select IdMovimiento, IdPeriodo, CodCuentaBanco, CuentaBanco, NroVoucher, "
            sSql1 += "NroItem, Fecha, CodCuentaFC, CuentaFC, '' as CuentaDestino, "
            sSql1 += "CodPersona, Persona, Glosa, Observaciones, "
            sSql1 += "CodArea, '' as Area, MontoBaseUSD, MontoIGVUSD, MontoUSD "
            sSql2 += "ORDER BY IdPeriodo, Fecha, NroVoucher"
            cn = New OleDbConnection()
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()
            cmd = New OleDbCommand(sSql1 + sSql2, cn)
            dtadap = New OleDbDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            gdvMovimiento.DataSource = dtaset.Tables(0)
            gdvMovimiento.PageIndex = pagina
            gdvMovimiento.DataBind()
            cn.Close()
        Else
            lblNumRegistros.Text = "No se encontro información"
            gdvMovimiento.DataSource = Nothing
            gdvMovimiento.DataBind()
        End If
    End Sub

    Protected Sub gdvMovimiento_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sSql As String
        Dim sIdPeriodo, sCodCuentaBanco, sNroVoucher, sNroItem, sObservaciones As String
        Dim sCodArea As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(0).Text = (e.Row.DataItemIndex + 1).ToString()
            sIdPeriodo = DataBinder.Eval(e.Row.DataItem, "IdPeriodo").ToString()
            sCodCuentaBanco = DataBinder.Eval(e.Row.DataItem, "CodCuentaBanco").ToString()
            sNroVoucher = DataBinder.Eval(e.Row.DataItem, "NroVoucher").ToString()
            sNroItem = DataBinder.Eval(e.Row.DataItem, "NroItem").ToString()
            sObservaciones = DataBinder.Eval(e.Row.DataItem, "Observaciones").ToString()
            sCodArea = DataBinder.Eval(e.Row.DataItem, "CodArea").ToString()
            If sCodArea <> "" Then
                sSql = "Select Area From Area Where CodArea = " + sCodArea
                e.Row.Cells(15).Text = Funciones.ExtraerValor(sSql, "Area")
            End If
            sSql = "Select cFC2.Cuenta2 as CuentaDestino "
            sSql += "From V_CuentaFC cFC2, Movimiento2 M2 "
            sSql += "Where cFC2.CodCuenta2 = M2.CodCuentaFC2 "
            sSql += "AND M2.IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND M2.CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND M2.NroVoucher = " + sNroVoucher + " "
            sSql += "AND M2.NroItem = " + sNroItem + " "
            If sObservaciones <> "" Then
                sSql += "AND M2.Observaciones = '" + sObservaciones + "' "
            End If
            Dim lnkCuentaDestino As Label = CType(e.Row.FindControl("lnkCuentaDestino"), Label)
            lnkCuentaDestino.Text = Funciones.ExtraerValor(sSql, "CuentaDestino")

            sSql = "Select a.Area as AreaNueva "
            sSql += "From Area a, Movimiento2 M2 "
            sSql += "Where a.CodArea = M2.CodArea2 "
            sSql += "AND M2.IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND M2.CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND M2.NroVoucher = " + sNroVoucher + " "
            sSql += "AND M2.NroItem = " + sNroItem + " "
            If sObservaciones <> "" Then
                sSql += "AND M2.Observaciones = '" + sObservaciones + "' "
            End If
            Dim lnkAreaNueva As Label = CType(e.Row.FindControl("lnkAreaNueva"), Label)
            lnkAreaNueva.Text = Funciones.ExtraerValor(sSql, "AreaNueva")
        End If
    End Sub

    Protected Sub gdvMovimiento_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        LlenaGridMovimiento(e.NewPageIndex)
        hdfPagina.Value = e.NewPageIndex
    End Sub

    Protected Sub gdvMovimiento_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gdvMovimiento.RowCommand
        If e.CommandName = "Editar" Then
            'Dim iRegistro As Int32 = Convert.ToInt32(e.CommandArgument) - (Convert.ToInt32(hdfPagina.Value) * Funciones.iREGxPAG)
            'Dim sidMovimiento As String = gdvMovimiento.DataKeys(iRegistro).Value
            CargarRegistro(e.CommandArgument)
            mpupDetalle.Show()
        End If
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click

    End Sub

    Private Sub CargarRegistro(ByVal sIdMovimiento As String)
        Dim sSql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtr As OleDbDataReader
        Dim sTipo As String = "", sEstado As String = ""

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        hdfIdMovimiento.Value = sIdMovimiento
        sSql = " Select IdMovimiento, IdPeriodo, CodCuentaBanco, CuentaBanco, NroVoucher, "
        sSql += "NroItem, Fecha, CodCuentaFC, CuentaFC, CodGrupoMov, GrupoMov, "
        sSql += "CodPersona, Persona, Glosa, Observaciones, "
        sSql += "CodArea, MontoBaseUSD, MontoIGVUSD, MontoUSD "
        sSql += "From Movimiento WHERE IdMovimiento = " + sIdMovimiento + ""
        cmd = New OleDbCommand(sSql, cn)
        dtr = cmd.ExecuteReader()
        If dtr.Read() Then
            lblIdPeriodo.Text = dtr("IdPeriodo").ToString()
            lblCodCuentaBanco.Text = dtr("CodCuentaBanco").ToString()
            lblCuentaBanco.Text = dtr("CuentaBanco").ToString()
            lblNroVoucher.Text = dtr("NroVoucher").ToString()
            lblNroItem.Text = dtr("NroItem").ToString()
            lblFecha.Text = String.Format("{0:dd/MM/yyyy}", dtr("Fecha"))
            lblCodCuentaFC.Text = dtr("CodCuentaFC").ToString()
            lblCuentaFC.Text = dtr("CuentaFC").ToString()
            lblCodGrupoMov.Text = dtr("CodGrupoMov").ToString()
            lblGrupoMov.Text = dtr("GrupoMov").ToString()
            lblCodPersona.Text = dtr("CodPersona").ToString()
            lblPersona.Text = dtr("Persona").ToString()
            lblGlosa.Text = dtr("Glosa").ToString()
            lblObservaciones.Text = dtr("Observaciones").ToString()
            lblCodArea.Text = dtr("CodArea").ToString()
            If lblCodArea.Text <> "" Then
                sSql = "Select Area From Area Where CodArea = " + lblCodArea.Text
                lblArea.Text = Funciones.ExtraerValor(sSql, "Area")
            End If
            lblMontoBaseUSD.Text = Funciones.FormatoDinero(dtr("MontoBaseUSD").ToString())
            lblMontoIGVUSD.Text = Funciones.FormatoDinero(dtr("MontoIGVUSD").ToString())
            lblMontoUSD.Text = Funciones.FormatoDinero(dtr("MontoUSD").ToString())

            LlenarCuentaFC2(cboCuentaFC2)
            sSql = "Select cFC2.CodCuenta2 "
            sSql += "From V_CuentaFC cFC2, Movimiento2 M2 "
            sSql += "Where cFC2.CodCuenta2 = M2.CodCuentaFC2 "
            sSql += "AND M2.IdPeriodo = " + lblIdPeriodo.Text + " "
            sSql += "AND M2.CodCuentaBanco = " + lblCodCuentaBanco.Text + " "
            sSql += "AND M2.NroVoucher = " + lblNroVoucher.Text + " "
            sSql += "AND M2.NroItem = " + lblNroItem.Text + " "
            If lblObservaciones.Text <> "" Then
                sSql += "AND M2.Observaciones = '" + lblObservaciones.Text + "' "
            End If
            Dim sCodCuenta2 As String = Funciones.ExtraerValor(sSql, "CodCuenta2")
            cboCuentaFC2.SelectedIndex = cboCuentaFC2.Items.IndexOf(cboCuentaFC2.Items.FindByValue(sCodCuenta2))

            LlenarArea2(cboArea2)
            sSql = "Select a.CodArea "
            sSql += "From Area a, Movimiento2 M2 "
            sSql += "Where a.CodArea = M2.CodArea2 "
            sSql += "AND M2.IdPeriodo = " + lblIdPeriodo.Text + " "
            sSql += "AND M2.CodCuentaBanco = " + lblCodCuentaBanco.Text + " "
            sSql += "AND M2.NroVoucher = " + lblNroVoucher.Text + " "
            sSql += "AND M2.NroItem = " + lblNroItem.Text + " "
            If lblObservaciones.Text <> "" Then
                sSql += "AND M2.Observaciones = '" + lblObservaciones.Text + "' "
            End If
            Dim sCodArea2 As String = Funciones.ExtraerValor(sSql, "CodArea")
            cboArea2.SelectedIndex = cboArea2.Items.IndexOf(cboArea2.Items.FindByValue(sCodArea2))
        End If
        cn.Close()
    End Sub

    Protected Sub btnCloseDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        mpupDetalle.Hide()
    End Sub

    Protected Sub btnSaveDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        Funciones.IngresarMovimiento2(hdfIdMovimiento.Value, cboCuentaFC2.SelectedValue, cboArea2.SelectedValue)
        LlenaGridMovimiento(hdfPagina.Value)
        mpupDetalle.Hide()
    End Sub

    Private Sub LlenarCuentaFC2(ByVal cboCuentaFC2 As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " select -1 as CodCuenta2, ' --SIN ASIGNAR-- ' AS Cuenta2 "
        sql += "UNION "
        sql += "SELECT CodCuenta2, Cuenta2 FROM CuentaFC2 ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCuentaFC2.DataSource = dtaset.Tables(0)
        cboCuentaFC2.DataValueField = "CodCuenta2"
        cboCuentaFC2.DataTextField = "Cuenta2"
        cboCuentaFC2.DataBind()
        cn.Close()
    End Sub

    Private Sub LlenarArea2(ByVal cboArea2 As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " select -1 as CodArea, ' --SIN ASIGNAR-- ' AS Area "
        sql += "UNION SELECT CodArea, Area FROM Area "
        sql += "ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboArea2.DataSource = dtaset.Tables(0)
        cboArea2.DataValueField = "CodArea"
        cboArea2.DataTextField = "Area"
        cboArea2.DataBind()
        cn.Close()
    End Sub

    Sub CargaResultadosAgrupadosExp()
        Dim sSql1, sSql2 As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        'Unimos los filtros
        sSql2 = "From V_Movimiento M, V_CuentaFC C, V_CuentaFC C2, Area A, Area A2, CuentaBanco B, Persona P "
        sSql2 &= "Where M.CodCuentaFC = C.CodCuenta2 and M.CodCuentaFC2 = C2.CodCuenta2 and M.CodArea = A.CodArea and M.CodArea2 = A2.CodArea "
        sSql2 &= "And M.CodPersona = P.CodPersona and M.CodCuentaBanco = B.CodCuenta and IdMovimiento > 0 "
        If cboPeriodo.Text <> "0" Then
            sSql2 += " AND IdPeriodo = " + cboPeriodo.Text + " "
        End If
        If cboCuentaBanco.Text <> "0" Then
            sSql2 += " AND CodCuentaBanco = " + cboCuentaBanco.Text + " "
        End If
        If cboCuentaFC.Text <> "0" Then
            sSql2 += " AND M.CodCuentaFC = " + cboCuentaFC.Text + " "
        End If
        If cboCuentaFCNueva.Text = "-1" Then
            sSql2 += " AND M.CodCuentaFC = M.CodCuentaFC2 "
        ElseIf cboCuentaFCNueva.Text = "-2" Then
            sSql2 += " AND M.CodCuentaFC <> M.CodCuentaFC2 "
        ElseIf cboCuentaFCNueva.Text <> "0" Then
            sSql2 += " AND M.CodCuentaFC2 = " + cboCuentaFC.Text + " "
        End If
        If txtPersona.Text <> "" Then
            sSql2 += " AND Persona like '%" + txtPersona.Text + "%' "
        End If
        If cboPersona.Text <> "0" Then
            sSql2 += " AND M.CodPersona = " + cboPersona.Text + " "
        End If
        If cboArea.Text <> "0" Then
            sSql2 += " AND M.CodArea = " + cboArea.Text + " "
        End If
        If cboAreaNueva.Text = "-1" Then
            sSql2 += " AND M.CodArea = M.CodArea2 "
        ElseIf cboAreaNueva.Text <> "0" Then
            sSql2 += " AND M.CodArea <> M.CodArea2 "
        End If
        If txtFecha.Text <> "" Then
            sSql2 += " AND Fecha = '" + txtFecha.Text + "' "
        End If
        If txtNroVoucher.Text <> "" Then
            sSql2 += " AND NroVoucher = " + txtNroVoucher.Text + " "
        End If
        If txtGlosa.Text <> "" Then
            sSql2 += " AND Glosa like '%" + txtGlosa.Text + "%' "
        End If

        'Extraemos el numero de registros
        hdfNroRegistros.Value = Funciones.ExtraerValor("Select Count(*) as NumReg " + sSql2, "NumReg")
        If hdfNroRegistros.Value <> "0" Then
            lblNumRegistros.Text = "Se han encontrado " + hdfNroRegistros.Value + " registros"
            'LLenamos la grila
            sSql1 = "Select IdPeriodo, M.CodCuentaBanco, Cuenta as CuentaBanco, NroVoucher, NroItem, Fecha, "
            sSql1 &= "M.CodCuentaFC, C.Cuenta2 as CuentaFC, case when M.CodCuentaFC <> M.CodCuentaFC2 then C2.Cuenta2 else null end as CuentaFC2, "
            sSql1 &= "M.CodPersona, Persona, Glosa, Observaciones, M.CodArea, A.Area, case when M.CodArea <> M.CodArea2 then A2.Area else null end as Area2, "
            sSql1 &= "CodGrupoMov, GrupoMov, MontoBaseUSD, MontoIGVUSD, MontoUSD "
            sSql2 &= "order by M.IdPeriodo, M.CodCuentaBanco, M.NroVoucher"
            cn = New OleDbConnection()
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()
            cmd = New OleDbCommand(sSql1 + sSql2, cn)
            dtadap = New OleDbDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            gdvResultadosExp.DataSource = dtaset.Tables(0)
            gdvResultadosExp.DataBind()
            cn.Close()
        Else
            lblNumRegistros.Text = "No se encontro información"
            gdvResultadosExp.DataSource = Nothing
            gdvResultadosExp.DataBind()
        End If
    End Sub

    Protected Sub gdvResultadosExp_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs)

        CargaResultadosAgrupadosExp()
        gdvResultadosExp.Visible = True

        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()
        gdvResultadosExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvResultadosExp)
        frm.RenderControl(htw)

        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=FlujoCaja.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
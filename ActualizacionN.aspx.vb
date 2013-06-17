Imports System.Data.OleDb
Imports System.Configuration

Public Class ActualizacionN
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ActualizacionN.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargarCombos()
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
        sql += "select 0 as IdPeriodo, ' --TODO-- ' as Periodo "
        sql += "UNION "
        sql += "SELECT IdPeriodo, Periodo FROM Periodo WHERE IdPeriodo >= 201101 "
        sql += "ORDER BY 1 "

        sql += "select 0 as CodCuenta, ' --TODAS LAS CUENTAS-- ' AS Cuenta "
        sql += "UNION SELECT CodCuenta, '[' + convert(varchar(10), CodCuenta) + '] - ' + Cuenta as Cuenta FROM CuentaBanco "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodCuenta, '-- TODAS LAS CUENTAS-- ' AS Cuenta "
        sql += "UNION SELECT CodCuenta, Cuenta + ' - [' + convert(varchar(10), CodCuenta) + ']' as Cuenta FROM CuentaFC "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodCuentaN, ' -- TODAS LAS CUENTAS -- ' AS CuentaN "
        sql += "UNION "
        sql += "select -2 as CodCuenta2, ' --ASIGNADOS-- ' AS CuentaFC2 "
        sql += "UNION "
        sql += "select -1 as CodCuentaN, ' --SIN ASIGNAR-- ' AS CuentaN "
        sql += "UNION SELECT CodCuentaN, CuentaN + ' - [' + convert(varchar(10), CodCuentaN) + ']' as CuentaN FROM CuentaFCN "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodPersona, ' --TODAS LAS PERSONAS-- ' AS Persona "
        sql += "UNION SELECT CodPersona, Persona FROM Persona "
        sql += "ORDER BY 2 "

        sql += "select 0 as CodArea, ' -- TODAS LAS AREAS -- ' AS Area "
        sql += "UNION "
        sql += "select -1 as CodArea, ' --SIN ASIGNAR-- ' AS Area "
        sql += "UNION SELECT CodArea, Area + ' - [' + convert(varchar(10), CodArea) + ']' as Area FROM Area "
        sql += "ORDER BY 2 "

        'sql += "select 0 as CodArea, ' --TODAS LAS AREAS-- ' AS Area "
        'sql += "UNION SELECT a.CodArea, a.Area FROM Area a, MovimientoN m Where a.CodArea = m.CodAreaN "
        'sql += "ORDER BY 2 "

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

        cboCuentaDestino.DataSource = dtaset.Tables(3)
        cboCuentaDestino.DataValueField = "CodCuentaN"
        cboCuentaDestino.DataTextField = "CuentaN"
        cboCuentaDestino.DataBind()

        cboPersona.DataSource = dtaset.Tables(4)
        cboPersona.DataValueField = "CodPersona"
        cboPersona.DataTextField = "Persona"
        cboPersona.DataBind()

        cboArea.DataSource = dtaset.Tables(5)
        cboArea.DataValueField = "CodArea"
        cboArea.DataTextField = "Area"
        cboArea.DataBind()

        cboAreaNueva.DataSource = dtaset.Tables(5)
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
        If cboCuentaDestino.Text = "-1" Then
            sSql2 += " AND IdMovimiento NOT IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones) "
        ElseIf cboCuentaDestino.Text = "-2" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones) "
        ElseIf cboCuentaDestino.Text <> "0" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones AND MN.CodCuentaFCN = " + cboCuentaDestino.Text + ") "
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
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones AND MN.CodAreaN = " + cboAreaNueva.Text + ") "
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
        Dim sIdPeriodo, sCodCuentaBanco, sNroVoucher, sNroItem, sObservaciones, sCodCuentaN As String
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

            sSql = "Select cFCN.CuentaN as CuentaDestino "
            sSql += "From V_CuentaFCN cFCN, MovimientoN MN "
            sSql += "Where cFCN.CodCuentaN = MN.CodCuentaFCN "
            sSql += "AND MN.IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND MN.CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND MN.NroVoucher = " + sNroVoucher + " "
            sSql += "AND MN.NroItem = " + sNroItem + " "
            If sObservaciones <> "" Then
                sSql += "AND MN.Observaciones = '" + sObservaciones + "' "
            End If
            Dim lnkCuentaDestino As Label = CType(e.Row.FindControl("lnkCuentaDestino"), Label)
            lnkCuentaDestino.Text = Funciones.ExtraerValor(sSql, "CuentaDestino")

            sSql = "Select a.Area as AreaNueva "
            sSql += "From Area a, MovimientoN MN "
            sSql += "Where a.CodArea = MN.CodAreaN "
            sSql += "AND MN.IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND MN.CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND MN.NroVoucher = " + sNroVoucher + " "
            sSql += "AND MN.NroItem = " + sNroItem + " "
            If sObservaciones <> "" Then
                sSql += "AND MN.Observaciones = '" + sObservaciones + "' "
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

            LlenarCuentaFCN(cboCuentaFCN)
            sSql = "Select cFCN.CodCuentaN "
            sSql += "From V_CuentaFCN cFCN, MovimientoN MN "
            sSql += "Where cFCN.CodCuentaN = MN.CodCuentaFCN "
            sSql += "AND MN.IdPeriodo = " + lblIdPeriodo.Text + " "
            sSql += "AND MN.CodCuentaBanco = " + lblCodCuentaBanco.Text + " "
            sSql += "AND MN.NroVoucher = " + lblNroVoucher.Text + " "
            sSql += "AND MN.NroItem = " + lblNroItem.Text + " "
            If lblObservaciones.Text <> "" Then
                sSql += "AND MN.Observaciones = '" + lblObservaciones.Text + "' "
            End If
            Dim sCodCuentaN As String = Funciones.ExtraerValor(sSql, "CodCuentaN")
            cboCuentaFCN.SelectedIndex = cboCuentaFCN.Items.IndexOf(cboCuentaFCN.Items.FindByValue(sCodCuentaN))

            LlenarAreaN(cboAreaN)
            sSql = "Select a.CodArea "
            sSql += "From Area a, MovimientoN MN "
            sSql += "Where a.CodArea = MN.CodAreaN "
            sSql += "AND MN.IdPeriodo = " + lblIdPeriodo.Text + " "
            sSql += "AND MN.CodCuentaBanco = " + lblCodCuentaBanco.Text + " "
            sSql += "AND MN.NroVoucher = " + lblNroVoucher.Text + " "
            sSql += "AND MN.NroItem = " + lblNroItem.Text + " "
            If lblObservaciones.Text <> "" Then
                sSql += "AND MN.Observaciones = '" + lblObservaciones.Text + "' "
            End If
            Dim sCodAreaN As String = Funciones.ExtraerValor(sSql, "CodArea")
            cboAreaN.SelectedIndex = cboAreaN.Items.IndexOf(cboAreaN.Items.FindByValue(sCodAreaN))
        End If
        cn.Close()
    End Sub

    Protected Sub btnCloseDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        mpupDetalle.Hide()
    End Sub

    Protected Sub btnSaveDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        Funciones.IngresarMovimientoN(hdfIdMovimiento.Value, cboCuentaFCN.SelectedValue, cboAreaN.SelectedValue)
        LlenaGridMovimiento(hdfPagina.Value)
        mpupDetalle.Hide()
    End Sub

    Private Sub LlenarCuentaFCN(ByVal cboCuentaFC2 As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " select -1 as CodCuentaN, ' --SIN ASIGNAR-- ' AS CuentaN "
        sql += "UNION "
        sql += "SELECT CodCuentaN, CuentaN FROM CuentaFCN ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCuentaFC2.DataSource = dtaset.Tables(0)
        cboCuentaFC2.DataValueField = "CodCuentaN"
        cboCuentaFC2.DataTextField = "CuentaN"
        cboCuentaFC2.DataBind()
        cn.Close()
    End Sub

    Private Sub LlenarAreaN(ByVal cboAreaN As DropDownList)
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

        cboAreaN.DataSource = dtaset.Tables(0)
        cboAreaN.DataValueField = "CodArea"
        cboAreaN.DataTextField = "Area"
        cboAreaN.DataBind()
        cn.Close()
    End Sub

    Sub CargaResultadosAgrupadosExp()
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
        If cboCuentaDestino.Text = "-1" Then
            sSql2 += " AND IdMovimiento NOT IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones) "
        ElseIf cboCuentaDestino.Text = "-2" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones) "
        ElseIf cboCuentaDestino.Text <> "0" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones AND MN.CodCuentaFCN = " + cboCuentaDestino.Text + ") "
        End If
        If cboPersona.Text <> "0" Then
            sSql2 += " AND CodPersona = " + cboPersona.Text + " "
        End If
        If txtPersona.Text <> "" Then
            sSql2 += " AND Persona like '%" + txtPersona.Text + "%' "
        End If
        If cboArea.Text = "-1" Then
            sSql2 += " AND CodArea IS NULL "
        ElseIf cboArea.Text <> "0" Then
            sSql2 += " AND CodArea = " + cboArea.Text + " "
        End If
        If cboAreaNueva.Text <> "0" Then
            sSql2 += " AND IdMovimiento IN (SELECT IdMovimiento FROM MovimientoN MN, Movimiento M WHERE MN.IdPeriodo = M.IdPeriodo AND MN.CodCuentaBanco = M.CodCuentaBanco AND MN.NroVoucher = M.NroVoucher AND MN.NroItem = M.NroItem AND MN.CodCuentaFC = M.CodCuentaFC AND MN.Observaciones = M.Observaciones AND MN.CodAreaN = " + cboAreaNueva.Text + ") "
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

            sSql = "Select cFCN.CuentaN as CuentaDestino "
                sSql += "From V_CuentaFCN cFCN, MovimientoN MN "
                sSql += "Where cFCN.CodCuentaN = MN.CodCuentaFCN "
                sSql += "AND MN.IdPeriodo = " + sIdPeriodo + " "
                sSql += "AND MN.CodCuentaBanco = " + sCodCuentaBanco + " "
                sSql += "AND MN.NroVoucher = " + sNroVoucher + " "
                sSql += "AND MN.NroItem = " + sNroItem + " "
                If sObservaciones <> "" Then
                    sSql += "AND MN.Observaciones = '" + sObservaciones + "' "
                End If
                Dim lnkCuentaDestino As Label = CType(e.Row.FindControl("lnkCuentaDestino"), Label)
                lnkCuentaDestino.Text = Funciones.ExtraerValor(sSql, "CuentaDestino")

                sSql = "Select a.Area as AreaNueva "
                sSql += "From Area a, MovimientoN MN "
                sSql += "Where a.CodArea = MN.CodAreaN "
                sSql += "AND MN.IdPeriodo = " + sIdPeriodo + " "
                sSql += "AND MN.CodCuentaBanco = " + sCodCuentaBanco + " "
                sSql += "AND MN.NroVoucher = " + sNroVoucher + " "
                sSql += "AND MN.NroItem = " + sNroItem + " "
                If sObservaciones <> "" Then
                    sSql += "AND MN.Observaciones = '" + sObservaciones + "' "
                End If
                Dim lnkAreaNueva As Label = CType(e.Row.FindControl("lnkAreaNueva"), Label)
                lnkAreaNueva.Text = Funciones.ExtraerValor(sSql, "AreaNueva")

            sCodArea = DataBinder.Eval(e.Row.DataItem, "CodArea").ToString()
            If sCodArea <> "" Then
                sSql = "Select Area From Area Where CodArea = " + sCodArea
                e.Row.Cells(15).Text = Funciones.ExtraerValor(sSql, "Area")
            End If
        End If
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
        Response.AddHeader("Content-Disposition", "attachment;filename=FlujoCajaNueva.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
Imports System.Data.OleDb
Imports System.Configuration

Public Class AjustesExt
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "AjustesExt.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargarCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
            gdvAjusteExt.PageSize = Funciones.iREGxPAG
            LlenaGridAjuste(0)
            lblMensaje.Text = ""
        End If
        If hdfAccion.Value = "Recargar" Then
            LlenaGridAjuste(0)
            hdfAccion.Value = ""
            lblMensaje.Text = ""
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
        sql += "SELECT IdPeriodo, Periodo FROM Periodo WHERE IdPeriodo >= 201001 "
        sql += "ORDER BY 1 "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboPeriodo.DataSource = dtaset.Tables(0)
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataBind()

        cn.Close()
    End Sub

    Sub LlenaGridAjuste(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdAjuste, IdPeriodo, Fecha, CodTipoAjuste, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes, Observacion, MontoUSD "
        sql += "From AjusteExt Where IdAjuste > 0"
        If cboPeriodo.Text <> "0" Then
            sql += " AND IdPeriodo = " + cboPeriodo.Text + " "
        End If
        sql += "ORDER BY IdPeriodo, Fecha, CodTipoAjuste desc, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes, Observacion"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvAjusteExt.DataSource = dtaset.Tables(0)
        gdvAjusteExt.PageIndex = pagina
        gdvAjusteExt.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvAjusteExt_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then

            Dim sCodTipoAjuste As String = DataBinder.Eval(e.Row.DataItem, "CodTipoAjuste").ToString()
            Dim lblCodTipoAjuste As Label = CType(e.Row.FindControl("lblCodTipoAjuste"), Label)
            If (sCodTipoAjuste = "I") Then
                lblCodTipoAjuste.Text = "Adicion"
            ElseIf (sCodTipoAjuste = "E") Then
                lblCodTipoAjuste.Text = "Cambio"
            Else
                lblCodTipoAjuste.Text = "Desconocido"
            End If

            Dim sCodCuentaFCOrig As String = DataBinder.Eval(e.Row.DataItem, "CodCuentaFCOrig").ToString()
            Dim lblCuentaFCOrig As Label = CType(e.Row.FindControl("lblCuentaFCOrig"), Label)
            lblCuentaFCOrig.Text = Funciones.ExtraerValor("Select Cuenta2 From V_CuentaFC Where CodCuenta2 = " + sCodCuentaFCOrig, "Cuenta2")

            Dim sCodPersonaOrig As String = DataBinder.Eval(e.Row.DataItem, "CodPersonaOrig").ToString()
            Dim lblPersonaOrig As Label = CType(e.Row.FindControl("lblPersonaOrig"), Label)
            lblPersonaOrig.Text = Funciones.ExtraerValor("Select Persona From Persona Where CodPersona = " + sCodPersonaOrig, "Persona")

            Dim sCodCuentaFCDes As String = DataBinder.Eval(e.Row.DataItem, "CodCuentaFCDes").ToString()
            Dim lblCuentaFCDes As Label = CType(e.Row.FindControl("lblCuentaFCDes"), Label)
            lblCuentaFCDes.Text = Funciones.ExtraerValor("Select Cuenta2 From V_CuentaFC Where CodCuenta2 = " + sCodCuentaFCDes, "Cuenta2")

            Dim sCodPersonaDes As String = DataBinder.Eval(e.Row.DataItem, "CodPersonaDes").ToString()
            Dim lblPersonaDes As Label = CType(e.Row.FindControl("lblPersonaDes"), Label)
            lblPersonaDes.Text = Funciones.ExtraerValor("Select Persona From Persona Where CodPersona = " + sCodPersonaDes, "Persona")

            Dim sCodAreaDes As String = DataBinder.Eval(e.Row.DataItem, "CodAreaDes").ToString()
            Dim lblAreaDes As Label = CType(e.Row.FindControl("lblAreaDes"), Label)
            lblAreaDes.Text = Funciones.ExtraerValor("Select Area From Area Where CodArea = " + sCodAreaDes, "Area")

        End If
    End Sub

    Protected Sub gdvAjusteExt_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdAjuste As String = gdvAjusteExt.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarAjusteExt(sIdAjuste)
        LlenaGridAjuste(0)
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        hdfAccion.Value = "Nuevo"
        'LlenarIdPeriodo(cboIdPeriodo)
        LlenarCodCuentaFC(cboCodCuentaFCOrig)
        LlenarCodPersona(cboCodPersonaOrig)
        LlenarCodCuentaFC(cboCodCuentaFCDes)
        LlenarCodPersona(cboCodPersonaDes)
        LlenarCodArea(cboCodAreaDes)
        txtFecha.Text = Date.Now.Date.ToString("yyyy-MM-dd")
        txtObservacion.Text = ""
        txtMontoUSD.Text = ""
        mpupDetalle.Show()
    End Sub

    Protected Sub gdvMovimiento_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gdvAjusteExt.RowCommand
        If e.CommandName = "Editar" Then
            hdfAccion.Value = "Editar"
            CargarRegistro(e.CommandArgument)
            mpupDetalle.Show()
        End If
    End Sub

    Private Sub CargarRegistro(ByVal sIdAjuste As String)
        Dim sSql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtr As OleDbDataReader
        Dim sTipo As String = "", sEstado As String = "", sCodTipoAjuste As String = ""

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        hdfIdAjuste.Value = sIdAjuste
        sSql = " Select IdAjuste, IdPeriodo, Fecha, CodTipoAjuste, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes, Observacion, MontoUSD "
        sSql += "From AjusteExt Where IdAjuste = " + sIdAjuste + ""
        cmd = New OleDbCommand(sSql, cn)
        dtr = cmd.ExecuteReader()
        If dtr.Read() Then
            'LlenarIdPeriodo(cboIdPeriodo)
            'cboIdPeriodo.SelectedValue = dtr("IdPeriodo").ToString()
            sCodTipoAjuste = dtr("CodTipoAjuste").ToString()
            If sCodTipoAjuste = "I" Then
                rdbIngreso.Checked = True
            ElseIf sCodTipoAjuste = "E" Then
                rdbEgreso.Checked = True
            End If
            If dtr("Fecha").ToString() <> "" Then
                txtFecha.Text = Convert.ToDateTime(dtr("Fecha")).ToString("yyyy-MM-dd")
            Else
                txtFecha.Text = ""
            End If
            LlenarCodCuentaFC(cboCodCuentaFCOrig)
            cboCodCuentaFCOrig.SelectedValue = dtr("CodCuentaFCOrig").ToString()
            LlenarCodPersona(cboCodPersonaOrig)
            cboCodPersonaOrig.SelectedValue = dtr("CodPersonaOrig").ToString()
            LlenarCodCuentaFC(cboCodCuentaFCDes)
            cboCodCuentaFCDes.SelectedValue = dtr("CodCuentaFCDes").ToString()
            LlenarCodPersona(cboCodPersonaDes)
            cboCodPersonaDes.SelectedValue = dtr("CodPersonaDes").ToString()
            LlenarCodArea(cboCodAreaDes)
            cboCodAreaDes.SelectedValue = dtr("CodAreaDes").ToString()
            txtObservacion.Text = dtr("Observacion").ToString()
            txtMontoUSD.Text = dtr("MontoUSD").ToString()
        End If
        cn.Close()
    End Sub

    Protected Sub btnCloseDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        mpupDetalle.Hide()
    End Sub

    Protected Sub btnSaveDetalle_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sCodTipoAjuste As String = "", sIdPeriodo As String = ""
        If rdbIngreso.Checked Then
            sCodTipoAjuste = "I"
        ElseIf rdbEgreso.Checked Then
            sCodTipoAjuste = "E"
        End If
        sIdPeriodo = txtFecha.Text.Substring(0, 4) & txtFecha.Text.Substring(5, 2)

        If hdfAccion.Value = "Editar" Then
            lblMensaje.Text = Funciones.ActualizarAjusteExt(hdfIdAjuste.Value, sIdPeriodo, txtFecha.Text, sCodTipoAjuste, cboCodCuentaFCOrig.SelectedValue, cboCodPersonaOrig.SelectedValue, cboCodCuentaFCDes.SelectedValue, cboCodPersonaDes.SelectedValue, cboCodAreaDes.SelectedValue, txtObservacion.Text, txtMontoUSD.Text)
        ElseIf hdfAccion.Value = "Nuevo" Then
            lblMensaje.Text = Funciones.InsertarAjusteExt(sIdPeriodo, txtFecha.Text, sCodTipoAjuste, cboCodCuentaFCOrig.SelectedValue, cboCodPersonaOrig.SelectedValue, cboCodCuentaFCDes.SelectedValue, cboCodPersonaDes.SelectedValue, cboCodAreaDes.SelectedValue, txtObservacion.Text, txtMontoUSD.Text)
        End If
        LlenaGridAjuste(0)
        mpupDetalle.Hide()
    End Sub

    Private Sub LlenarIdPeriodo(ByVal cboIdPeriodo As DropDownList)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "SELECT IdPeriodo, Periodo FROM Periodo "
        sql += "ORDER BY 1 "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboIdPeriodo.DataSource = dtaset.Tables(0)
        cboIdPeriodo.DataValueField = "IdPeriodo"
        cboIdPeriodo.DataTextField = "Periodo"
        cboIdPeriodo.DataBind()
        cn.Close()
    End Sub

    Private Sub LlenarCodCuentaFC(ByVal cboCodCuentaFC As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodCuenta2, Cuenta2 FROM V_CuentaFC ORDER BY Cuenta2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCuentaFC.DataSource = dtaset.Tables(0)
        cboCodCuentaFC.DataValueField = "CodCuenta2"
        cboCodCuentaFC.DataTextField = "Cuenta2"
        cboCodCuentaFC.DataBind()
        cn.Close()
    End Sub

    Private Sub LlenarCodPersona(ByVal cboCodPersona As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodPersona, Persona FROM Persona ORDER BY Persona"
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodPersona.DataSource = dtaset.Tables(0)
        cboCodPersona.DataValueField = "CodPersona"
        cboCodPersona.DataTextField = "Persona"
        cboCodPersona.DataBind()
        cn.Close()
    End Sub

    Private Sub LlenarCodArea(ByVal cboCodArea As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT 0 as CodArea, '-- Sin Area --' as Area " & _
                "union select CodArea, Area " & _
                "FROM Area " & _
                "WHERE isnull(FlagPrograma, '') = 'S' " & _
                "ORDER BY Area"
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodArea.DataSource = dtaset.Tables(0)
        cboCodArea.DataValueField = "CodArea"
        cboCodArea.DataTextField = "Area"
        cboCodArea.DataBind()
        cn.Close()
    End Sub


    Sub CargaResultadosAgrupadosExp()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdAjuste, IdPeriodo, Fecha, CodTipoAjuste, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes,Observacion, MontoUSD "
        sql += "From AjusteExt Where IdAjuste > 0"
        If cboPeriodo.Text <> "0" Then
            sql += " AND IdPeriodo = " + cboPeriodo.Text + " "
        End If
        sql += "ORDER BY IdPeriodo, Fecha, CodTipoAjuste desc, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes, Observacion"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvResultadosExp.DataSource = dtaset.Tables(0)
        gdvResultadosExp.DataBind()
        cn.Close()
    End Sub

    Protected Sub gdvResultadosExp_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then

            Dim sCodTipoAjuste As String = DataBinder.Eval(e.Row.DataItem, "CodTipoAjuste").ToString()
            Dim lblCodTipoAjuste As Label = CType(e.Row.FindControl("lblCodTipoAjuste"), Label)
            If (sCodTipoAjuste = "I") Then
                lblCodTipoAjuste.Text = "Adicion"
            ElseIf (sCodTipoAjuste = "E") Then
                lblCodTipoAjuste.Text = "Cambio"
            Else
                lblCodTipoAjuste.Text = "Desconocido"
            End If

            Dim sCodCuentaFCOrig As String = DataBinder.Eval(e.Row.DataItem, "CodCuentaFCOrig").ToString()
            Dim lblCuentaFCOrig As Label = CType(e.Row.FindControl("lblCuentaFCOrig"), Label)
            lblCuentaFCOrig.Text = Funciones.ExtraerValor("Select Cuenta2 From V_CuentaFC Where CodCuenta2 = " + sCodCuentaFCOrig, "Cuenta2")

            Dim sCodPersonaOrig As String = DataBinder.Eval(e.Row.DataItem, "CodPersonaOrig").ToString()
            Dim lblPersonaOrig As Label = CType(e.Row.FindControl("lblPersonaOrig"), Label)
            lblPersonaOrig.Text = Funciones.ExtraerValor("Select Persona From Persona Where CodPersona = " + sCodPersonaOrig, "Persona")

            Dim sCodCuentaFCDes As String = DataBinder.Eval(e.Row.DataItem, "CodCuentaFCDes").ToString()
            Dim lblCuentaFCDes As Label = CType(e.Row.FindControl("lblCuentaFCDes"), Label)
            lblCuentaFCDes.Text = Funciones.ExtraerValor("Select Cuenta2 From V_CuentaFC Where CodCuenta2 = " + sCodCuentaFCDes, "Cuenta2")

            Dim sCodPersonaDes As String = DataBinder.Eval(e.Row.DataItem, "CodPersonaDes").ToString()
            Dim lblPersonaDes As Label = CType(e.Row.FindControl("lblPersonaDes"), Label)
            lblPersonaDes.Text = Funciones.ExtraerValor("Select Persona From Persona Where CodPersona = " + sCodPersonaDes, "Persona")

            Dim sCodAreaDes As String = DataBinder.Eval(e.Row.DataItem, "CodAreaDes").ToString()
            Dim lblAreaDes As Label = CType(e.Row.FindControl("lblAreaDes"), Label)
            lblAreaDes.Text = Funciones.ExtraerValor("Select Area From Area Where CodArea = " + sCodAreaDes, "Area")
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
        Response.AddHeader("Content-Disposition", "attachment;filename=AjusteExterior.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub
End Class
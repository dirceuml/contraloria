Imports System.Data.OleDb
Imports System.Configuration

Public Class Ajustes
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Ajustes.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargarCombos()
            cboPeriodo.SelectedValue = Date.Now.ToString("yyyyMM")
            cboCuentaFC2.SelectedValue = "(TESORERIA)"
            gdvAjuste.PageSize = Funciones.iREGxPAG
            LlenaGridAjuste(0)
            lblMensaje.Text = ""
        End If
        If hdfAccion.Value = "Recargar" Then
            LlenaGridAjuste(0)
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
        sql += "SELECT IdPeriodo, Periodo FROM Periodo "
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

        sql = " Select IdAjuste, IdPeriodo, Fecha, CodCuentaFC2, MontoUSD "
        sql += "From Ajuste Where IdAjuste > 0 "
        If cboPeriodo.Text <> "0" Then
            sql += " AND IdPeriodo = " + cboPeriodo.Text + " "
        End If
        If cboCuentaFC2.Text <> "" Then
            sql += " AND CodCuentaFC2 IN (Select CodCuenta2 From CuentaFC2 Where Cuenta2 LIKE '%" + cboCuentaFC2.Text + "%') "
        End If
        sql += "ORDER BY IdPeriodo, Fecha, CodCuentaFC2"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvAjuste.DataSource = dtaset.Tables(0)
        gdvAjuste.PageIndex = pagina
        gdvAjuste.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvAjuste_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                'Dim cboIdPeriodo As DropDownList = CType(e.Row.FindControl("cboIdPeriodo"), DropDownList)
                'LlenarIdPeriodo(cboIdPeriodo)
                'cboIdPeriodo.SelectedValue = DataBinder.Eval(e.Row.DataItem, "IdPeriodo").ToString()
                Dim sFecha As String = DataBinder.Eval(e.Row.DataItem, "Fecha").ToString()
                Dim txtFecha As TextBox = CType(e.Row.FindControl("txtFecha"), TextBox)
                If sFecha <> "" Then
                    txtFecha.Text = Convert.ToDateTime(DataBinder.Eval(e.Row.DataItem, "Fecha")).ToString("yyyy-MM-dd")
                Else
                    txtFecha.Text = ""
                End If

                Dim cboCodCuentaFC2 As DropDownList = CType(e.Row.FindControl("cboCodCuentaFC2"), DropDownList)
                LlenarCodCuentaFC2(cboCodCuentaFC2)
                cboCodCuentaFC2.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCuentaFC2").ToString()
            Else
                Dim sFecha As String = DataBinder.Eval(e.Row.DataItem, "Fecha").ToString()
                Dim lblFecha As Label = CType(e.Row.FindControl("lblFecha"), Label)
                If sFecha <> "" Then
                    lblFecha.Text = Convert.ToDateTime(sFecha).ToString("yyyy-MM-dd")
                Else
                    lblFecha.Text = ""
                End If
                Dim sCodCuentaFC2 As String = DataBinder.Eval(e.Row.DataItem, "CodCuentaFC2").ToString()
                Dim lblCuentaFC2 As Label = CType(e.Row.FindControl("lblCuentaFC2"), Label)
                lblCuentaFC2.Text = Funciones.ExtraerValor("Select Cuenta2 From CuentaFC2 Where CodCuenta2 = " + sCodCuentaFC2, "Cuenta2")
                Dim lnkEdit As LinkButton = CType(e.Row.FindControl("lnkEdit"), LinkButton)
                lnkEdit.Visible = (sCodCuentaFC2 <> "901" And sCodCuentaFC2 <> "928" And sCodCuentaFC2 <> "946")
                Dim lnkDelete As LinkButton = CType(e.Row.FindControl("lnkDelete"), LinkButton)
                lnkDelete.Visible = (sCodCuentaFC2 <> "901" And sCodCuentaFC2 <> "928" And sCodCuentaFC2 <> "946")
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            'Dim cboIdPeriodoNew As DropDownList = CType(e.Row.FindControl("cboIdPeriodoNew"), DropDownList)
            'LlenarIdPeriodo(cboIdPeriodoNew)
            Dim txtFechaNew As TextBox = CType(e.Row.FindControl("txtFechaNew"), TextBox)
            txtFechaNew.Text = Date.Now.Date.ToString("yyyy-MM-dd")

            Dim cboCodCuentaFC2New As DropDownList = CType(e.Row.FindControl("cboCodCuentaFC2New"), DropDownList)
            LlenarCodCuentaFC2(cboCodCuentaFC2New)
        End If
    End Sub

    Protected Sub gdvAjuste_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvAjuste.EditIndex = e.NewEditIndex
        LlenaGridAjuste(0)
    End Sub

    Protected Sub gdvAjuste_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvAjuste.Rows(e.RowIndex)
        Dim sIdAjuste As String = gdvAjuste.DataKeys(e.RowIndex).Value.ToString()
        'Dim cboIdPeriodo As DropDownList = DirectCast(row.FindControl("cboIdPeriodo"), DropDownList)
        Dim sIdPeriodo As String = ""
        Dim txtFecha As TextBox = DirectCast(row.FindControl("txtFecha"), TextBox)
        sIdPeriodo = txtFecha.Text.Substring(0, 4) & txtFecha.Text.Substring(5, 2)
        Dim cboCodCuentaFC2 As DropDownList = DirectCast(row.FindControl("cboCodCuentaFC2"), DropDownList)
        Dim sMontoUSD As TextBox = DirectCast(row.FindControl("txtMontoUSD"), TextBox)
        'Dim sMontoBaseUSD As TextBox = DirectCast(row.FindControl("txtMontoBaseUSD"), TextBox)
        'Dim sMontoIGVUSD As TextBox = DirectCast(row.FindControl("txtMontoIGVUSD"), TextBox)
        'Dim sMontoUSD As Decimal = Decimal.Parse(sMontoBaseUSD.Text) + Decimal.Parse(sMontoIGVUSD.Text)
        lblMensaje.Text = Funciones.ActualizarAjuste(sIdAjuste, sIdPeriodo, txtFecha.Text, cboCodCuentaFC2.SelectedValue, sMontoUSD.Text)
        gdvAjuste.EditIndex = -1
        LlenaGridAjuste(0)
    End Sub

    Protected Sub gdvAjuste_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdAjuste As String = gdvAjuste.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarAjuste(sIdAjuste)
        LlenaGridAjuste(0)
    End Sub

    Protected Sub gdvAjuste_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvAjuste.EditIndex = -1
        LlenaGridAjuste(0)
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Dim cboIdPeriodo As DropDownList = DirectCast(gdvAjuste.FooterRow.FindControl("cboIdPeriodoNew"), DropDownList)
        Dim sIdPeriodo As String = ""
        Dim txtFecha As TextBox = DirectCast(gdvAjuste.FooterRow.FindControl("txtFechaNew"), TextBox)
        sIdPeriodo = txtFecha.Text.Substring(0, 4) & txtFecha.Text.Substring(5, 2)
        Dim cboCodCuentaFC2 As DropDownList = DirectCast(gdvAjuste.FooterRow.FindControl("cboCodCuentaFC2New"), DropDownList)
        Dim sMontoUSD As TextBox = DirectCast(gdvAjuste.FooterRow.FindControl("txtMontoUSDNew"), TextBox)
        'Dim sMontoBaseUSD As TextBox = DirectCast(gdvAjuste.FooterRow.FindControl("txtMontoBaseUSDNew"), TextBox)
        'Dim sMontoIGVUSD As TextBox = DirectCast(gdvAjuste.FooterRow.FindControl("txtMontoIGVUSDNew"), TextBox)
        'Dim sMontoUSD As Decimal = Decimal.Parse(sMontoBaseUSD.Text) + Decimal.Parse(sMontoIGVUSD.Text)
        lblMensaje.Text = Funciones.InsertarAjuste(sIdPeriodo, txtFecha.Text, cboCodCuentaFC2.SelectedValue, sMontoUSD.Text)
        gdvAjuste.EditIndex = -1
        LlenaGridAjuste(0)
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

    Private Sub LlenarCodCuentaFC2(ByVal cboCodCuentaFC2 As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodCuenta2, Cuenta2 FROM CuentaFC2 WHERE Cuenta2 LIKE 'AJUSTE%' AND CodEstado > 0 ORDER BY Cuenta2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCuentaFC2.DataSource = dtaset.Tables(0)
        cboCodCuentaFC2.DataValueField = "CodCuenta2"
        cboCodCuentaFC2.DataTextField = "Cuenta2"
        cboCodCuentaFC2.DataBind()
        cn.Close()
    End Sub

    Sub CargaResultadosAgrupadosExp()
        Dim Sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        Sql = " Select A.IdAjuste, A.IdPeriodo, A.Fecha, C.Cuenta2, A.MontoUSD "
        Sql += "From Ajuste A, CuentaFC2 c Where A.CodCuentaFC2 = C.CodCuenta2 "
        If cboPeriodo.Text <> "0" Then
            Sql += " AND A.IdPeriodo = " + cboPeriodo.Text + " "
        End If
        If cboCuentaFC2.Text <> "" Then
            Sql += " AND C.Cuenta2 LIKE '%" + cboCuentaFC2.Text + "%' "
        End If
        Sql += "ORDER BY A.IdPeriodo, Fecha, C.Cuenta2"

        cmd = New OleDbCommand(Sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvResultadosExp.DataSource = dtaset.Tables(0)
        gdvResultadosExp.DataBind()
        cn.Close()
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
        Response.AddHeader("Content-Disposition", "attachment;filename=Ajuste.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
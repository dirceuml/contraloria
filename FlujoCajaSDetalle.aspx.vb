Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCajaSDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "FlujoCajaS.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("cod") = "" Then Response.Redirect("FlujoCajaS.aspx")

        If Not IsPostBack Then
            hdfCodDetalle.Value = Page.Request.QueryString("cod")
            lblCodDetalle.Text = hdfCodDetalle.Value
            gdvCtaDetalleS.PageSize = Funciones.iREGxPAG
            LlenaGridCtaDetalleS(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaDetalleS(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaDetalle, CodDetalle, Signo, CodCtaOrigen, TipoDetalle "
        sql += "From CtaDetalleS "
        sql += "Where CodDetalle = '" + hdfCodDetalle.Value + "' "
        sql += "ORDER BY 1"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaDetalleS.DataSource = dtaset.Tables(0)
        gdvCtaDetalleS.PageIndex = pagina
        gdvCtaDetalleS.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaDetalleS_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sSigno, sCodCtaOrigen, sTipoDetalle As String

        If e.Row.RowType = DataControlRowType.DataRow Then

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
                LlenarCodCtaOrigen(cboCodCtaOrigen)
                cboCodCtaOrigen.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
                Dim cboTipoDetalle As DropDownList = CType(e.Row.FindControl("cboTipoDetalle"), DropDownList)
                cboTipoDetalle.SelectedValue = DataBinder.Eval(e.Row.DataItem, "TipoDetalle").ToString()
            Else
                sSigno = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                If sSigno = "1" Then
                    sSigno = " + "
                ElseIf sSigno = "-1" Then
                    sSigno = " - "
                End If
                Dim lnkSigno As Label = CType(e.Row.FindControl("lnkSigno"), Label)
                lnkSigno.Text = sSigno

                sCodCtaOrigen = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
                Dim lblCodCtaOrigen As Label = CType(e.Row.FindControl("lblCodCtaOrigen"), Label)
                lblCodCtaOrigen.Text = Funciones.ExtraerValor("Select Cuenta2 + ' [' + convert(varchar, CodCuenta2) + ']' as Cuenta2 From V_CuentaFC Where CodCuenta2 = '" + sCodCtaOrigen + "'", "Cuenta2")

                sTipoDetalle = DataBinder.Eval(e.Row.DataItem, "TipoDetalle").ToString()
                If sTipoDetalle = "GRP" Then
                    sTipoDetalle = "Grupo"
                ElseIf sTipoDetalle = "PER" Then
                    sTipoDetalle = "Persona"
                Else
                    sTipoDetalle = " "
                End If
                Dim lblTipoDetalle As Label = CType(e.Row.FindControl("lblTipoDetalle"), Label)
                lblTipoDetalle.Text = sTipoDetalle
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigenNew"), DropDownList)
            LlenarCodCtaOrigen(cboCodCtaOrigen)
        End If
    End Sub

    Protected Sub gdvCtaDetalleS_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaDetalleS.EditIndex = e.NewEditIndex
        LlenaGridCtaDetalleS(0)
    End Sub

    Protected Sub gdvCtaDetalleS_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaDetalleS.Rows(e.RowIndex)
        Dim sIdCtaDetalle As String = gdvCtaDetalleS.DataKeys(e.RowIndex).Value.ToString()
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(row.FindControl("cboTipoDetalle"), DropDownList)
        lblMensaje.Text = Funciones.ActualizarCtaDetalleS(sIdCtaDetalle, hdfCodDetalle.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue)
        gdvCtaDetalleS.EditIndex = -1
        LlenaGridCtaDetalleS(0)
    End Sub

    Protected Sub gdvCtaDetalleS_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaDetalleS As String = gdvCtaDetalleS.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaDetalleS(sIdCtaDetalleS)
        LlenaGridCtaDetalleS(0)
    End Sub

    Protected Sub gdvCtaDetalleS_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaDetalleS.EditIndex = -1
        LlenaGridCtaDetalleS(0)
    End Sub

    Private Sub LlenarCodCtaOrigen(ByVal cboCodCtaOrigen As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodCuenta2, Cuenta2 + ' [' + convert(varchar, CodCuenta2) + ']' as Cuenta2 FROM V_CuentaFC ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCtaOrigen.DataSource = dtaset.Tables(0)
        cboCodCtaOrigen.DataValueField = "CodCuenta2"
        cboCodCtaOrigen.DataTextField = "Cuenta2"
        cboCodCtaOrigen.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaDetalleS.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaDetalleS.FooterRow.FindControl("cboCodCtaOrigenNew"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(gdvCtaDetalleS.FooterRow.FindControl("cboTipoDetalleNew"), DropDownList)
        lblMensaje.Text = Funciones.InsertarCtaDetalleS(hdfCodDetalle.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue)
        gdvCtaDetalleS.EditIndex = -1
        LlenaGridCtaDetalleS(0)
    End Sub

End Class
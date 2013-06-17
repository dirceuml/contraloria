
Public Class CtasPorCobrar
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "CtasPorCobrar.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
            gdvCtasPorCobrar.DataBind()
        End If
    End Sub

    Protected Sub lnkNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaPorCobrar As TextBox = DirectCast(gdvCtasPorCobrar.FooterRow.FindControl("txtCtaPorCobrar"), TextBox)
        Dim txtAbreviado As TextBox = DirectCast(gdvCtasPorCobrar.FooterRow.FindControl("txtAbreviado"), TextBox)
        Dim cboCuentaEGP As DropDownList = DirectCast(gdvCtasPorCobrar.FooterRow.FindControl("cboCuentaEGP"), DropDownList)
        Dim cboSeccion As DropDownList = DirectCast(gdvCtasPorCobrar.FooterRow.FindControl("cboSeccion"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtasPorCobrar.FooterRow.FindControl("txtOrden"), TextBox)

        FuncionesCtasPorCobrar.CreaCtaPorCobrar(txtCtaPorCobrar.Text, txtAbreviado.Text, cboCuentaEGP.SelectedValue, cboSeccion.SelectedValue, txtOrden.Text)

        gdvCtasPorCobrar.EditIndex = -1
        gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrar.DataBind()
    End Sub

    Protected Sub gdvCtasPorCobrar_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtasPorCobrar.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCuentaEGP As DropDownList = DirectCast(e.Row.FindControl("cboCuentaEGP"), DropDownList)
                cboCuentaEGP.DataSource = LlenaCuentasEGP_CtaPorCobrar()
                cboCuentaEGP.DataTextField = "Cuenta"
                cboCuentaEGP.DataValueField = "CodCuenta"
                cboCuentaEGP.DataBind()
                cboCuentaEGP.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCuenta").ToString()
            Else
                'nada
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCuentaEGP As DropDownList = DirectCast(e.Row.FindControl("cboCuentaEGP"), DropDownList)
            cboCuentaEGP.DataSource = LlenaCuentasEGP_CtaPorCobrar()
            cboCuentaEGP.DataTextField = "Cuenta"
            cboCuentaEGP.DataValueField = "CodCuenta"
            cboCuentaEGP.DataBind()
        End If
    End Sub

    Protected Sub gdvCtasPorCobrar_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtasPorCobrar.RowEditing
        gdvCtasPorCobrar.EditIndex = e.NewEditIndex
        gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrar.DataBind()
    End Sub

    Protected Sub gdvCtasPorCobrar_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtasPorCobrar.RowUpdating
        Dim IdCtaPorCobrar, CodCuenta As String
        Dim row As GridViewRow = gdvCtasPorCobrar.Rows(e.RowIndex)
        Dim txtCtaPorCobrar As TextBox = DirectCast(row.FindControl("txtCtaPorCobrar"), TextBox)
        Dim txtAbreviado As TextBox = DirectCast(row.FindControl("txtAbreviado"), TextBox)
        Dim cboCuentaEGP As DropDownList = DirectCast(row.FindControl("cboCuentaEGP"), DropDownList)
        Dim cboSeccion As DropDownList = DirectCast(row.FindControl("cboSeccion"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)

        IdCtaPorCobrar = gdvCtasPorCobrar.DataKeys(e.RowIndex).Values(0).ToString
        CodCuenta = gdvCtasPorCobrar.DataKeys(e.RowIndex).Values(1).ToString

        ActualizaCtaPorCobrar(IdCtaPorCobrar, CodCuenta, txtCtaPorCobrar.Text, txtAbreviado.Text, cboCuentaEGP.SelectedValue, cboSeccion.SelectedValue, txtOrden.Text)

        gdvCtasPorCobrar.EditIndex = -1
        gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrar.DataBind()
    End Sub

    Protected Sub gdvCtasPorCobrar_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtasPorCobrar.RowCancelingEdit
        gdvCtasPorCobrar.EditIndex = -1
        gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrar.DataBind()
    End Sub

    Protected Sub gdvCtasPorCobrar_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gdvCtasPorCobrar.RowDeleting
        Dim IdCtaPorCobrar, CodCuenta As String
        Dim row As GridViewRow = gdvCtasPorCobrar.Rows(e.RowIndex)

        IdCtaPorCobrar = gdvCtasPorCobrar.DataKeys(e.RowIndex).Values(0).ToString
        CodCuenta = gdvCtasPorCobrar.DataKeys(e.RowIndex).Values(1).ToString

        EliminaCtaPorCobrar(IdCtaPorCobrar, CodCuenta)

        gdvCtasPorCobrar.EditIndex = -1
        gdvCtasPorCobrar.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrar.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvCtasPorCobrarExp.DataSource = LlenaCtasPorCobrar()
        gdvCtasPorCobrarExp.DataBind()
        gdvCtasPorCobrarExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvCtasPorCobrarExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Ctas por Cobrar.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
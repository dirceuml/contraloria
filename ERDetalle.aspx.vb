
Public Class ERDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ER.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("IdCtaER") = "" Then Response.Redirect("ER.aspx")

        If Not IsPostBack Then
            hdfIdCtaER.Value = Request.QueryString("IdCtaER")
            lblCtaER.Text = FuncionesInfGestion.BuscaCtaER(hdfIdCtaER.Value)
            gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
            gdvCtaERDet.DataBind()
            If FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value).Rows(0)("IdCtaERDet").ToString = "0" Then gdvCtaERDet.Rows(0).Visible = False
        End If
    End Sub

    Protected Sub gdvCtaERDet_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaERDet.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()

                Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
                cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasER()
                cboCodCtaOrigen.DataBind()
                cboCodCtaOrigen.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
            cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasER()
            cboCodCtaOrigen.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaERDet_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaERDet.RowEditing
        gdvCtaERDet.EditIndex = e.NewEditIndex
        gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
        gdvCtaERDet.DataBind()
    End Sub

    Protected Sub gdvCtaERDet_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaERDet.RowUpdating
        Dim row As GridViewRow = gdvCtaERDet.Rows(e.RowIndex)
        Dim IdCtaERDet As String = gdvCtaERDet.DataKeys(e.RowIndex).Value.ToString()
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)

        FuncionesInfGestion.ActualizaCtaERDet(IdCtaERDet, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue)
        gdvCtaERDet.EditIndex = -1
        gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
        gdvCtaERDet.DataBind()
    End Sub

    Protected Sub gdvCtaERDet_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaERDet.RowDeleting
        Dim IdCtaERDet As String = gdvCtaERDet.DataKeys(e.RowIndex).Value.ToString()

        FuncionesInfGestion.EliminaCtaERDet(IdCtaERDet)
        gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
        gdvCtaERDet.DataBind()
        If FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value).Rows(0)("IdCtaERDet").ToString = "0" Then gdvCtaERDet.Rows(0).Visible = False
    End Sub

    Protected Sub gdvCtaERDet_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaERDet.RowCancelingEdit
        gdvCtaERDet.EditIndex = -1
        gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
        gdvCtaERDet.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaERDet.FooterRow.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaERDet.FooterRow.FindControl("cboCodCtaOrigen"), DropDownList)

        FuncionesInfGestion.AgregaCtaERDet(hdfIdCtaER.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue)
        gdvCtaERDet.EditIndex = -1
        gdvCtaERDet.DataSource = FuncionesInfGestion.LlenaERDet(hdfIdCtaER.Value)
        gdvCtaERDet.DataBind()
    End Sub

End Class
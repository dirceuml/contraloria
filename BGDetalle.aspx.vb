Imports System.Data.OleDb
Imports System.Configuration

Public Class BGDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "BG.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("IdCtaBG") = "" Then Response.Redirect("BG.aspx")

        If Not IsPostBack Then
            hdfIdCtaBG.Value = Page.Request.QueryString("IdCtaBG")
            lblCtaBG.Text = FuncionesInfGestion.BuscaCtaBG(hdfIdCtaBG.Value)
            gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
            gdvCtaBGDet.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaBGDet_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaBGDet.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
                cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasBG
                cboCodCtaOrigen.DataBind()
                cboCodCtaOrigen.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
            cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasBG
            cboCodCtaOrigen.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaBGDet_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaBGDet.RowEditing
        gdvCtaBGDet.EditIndex = e.NewEditIndex
        gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
        gdvCtaBGDet.DataBind()
    End Sub

    Protected Sub gdvCtaBGDet_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaBGDet.RowUpdating
        Dim IdCtaBGDet As String = gdvCtaBGDet.DataKeys(e.RowIndex).Value.ToString()
        Dim row As GridViewRow = gdvCtaBGDet.Rows(e.RowIndex)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)

        FuncionesInfGestion.ActualizaCtaBGDet(IdCtaBGDet, cboCodCtaOrigen.SelectedValue)
        gdvCtaBGDet.EditIndex = -1
        gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
        gdvCtaBGDet.DataBind()
    End Sub

    Protected Sub gdvCtaBGDet_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaBGDet.RowDeleting
        Dim IdCtaBGDet As String = gdvCtaBGDet.DataKeys(e.RowIndex).Value.ToString()

        FuncionesInfGestion.EliminaCtaBGDet(IdCtaBGDet)
        gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
        gdvCtaBGDet.DataBind()
    End Sub

    Protected Sub gdvCtaBGDet_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaBGDet.RowCancelingEdit
        gdvCtaBGDet.EditIndex = -1
        gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
        gdvCtaBGDet.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaBGDet.FooterRow.FindControl("cboCodCtaOrigen"), DropDownList)
        FuncionesInfGestion.AgregaCtaBGDet(hdfIdCtaBG.Value, cboCodCtaOrigen.SelectedValue)
        gdvCtaBGDet.EditIndex = -1
        gdvCtaBGDet.DataSource = FuncionesInfGestion.LlenaBGDet(hdfIdCtaBG.Value)
        gdvCtaBGDet.DataBind()
    End Sub

End Class
Imports System.Data.OleDb
Imports System.Configuration

Public Class EOAFDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "EOAF.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("IdCtaEOAF") = "" Then Response.Redirect("EOAF.aspx")

        If Not IsPostBack Then
            hdfIdCtaEOAF.Value = Page.Request.QueryString("IdCtaEOAF")
            lblCtaEOAF.Text = FuncionesInfGestion.BuscaCtaEOAF(hdfIdCtaEOAF.Value)
            gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
            gdvCtaEOAFDet.DataBind()
            If FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value).Rows(0)("IdCtaEOAFDet").ToString = "0" Then gdvCtaEOAFDet.Rows(0).Visible = False
        End If
    End Sub

    Protected Sub gdvCtaEOAFDet_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaEOAFDet.RowDataBound
        Dim Signo As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
                cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasFC()
                cboCodCtaOrigen.DataBind()
                cboCodCtaOrigen.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
            Else
                Signo = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                If Signo = "1" Then Signo = " + " Else Signo = " - "
                Dim lblSigno As Label = CType(e.Row.FindControl("lblSigno"), Label)
                lblSigno.Text = Signo
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
            cboCodCtaOrigen.DataSource = FuncionesInfGestion.LlenaCuentasFC()
            cboCodCtaOrigen.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaEOAFDet_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaEOAFDet.RowEditing
        gdvCtaEOAFDet.EditIndex = e.NewEditIndex
        gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
        gdvCtaEOAFDet.DataBind()
    End Sub

    Protected Sub gdvCtaEOAFDet_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaEOAFDet.RowUpdating
        Dim IdCtaEOAFDet As String = gdvCtaEOAFDet.DataKeys(e.RowIndex).Value.ToString()
        Dim row As GridViewRow = gdvCtaEOAFDet.Rows(e.RowIndex)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)

        FuncionesInfGestion.ActualizaCtaEOAFDet(IdCtaEOAFDet, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue)
        gdvCtaEOAFDet.EditIndex = -1
        gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
        gdvCtaEOAFDet.DataBind()
    End Sub

    Protected Sub gdvCtaEOAFDet_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaEOAFDet.RowDeleting
        Dim IdCtaEOAFDet As String = gdvCtaEOAFDet.DataKeys(e.RowIndex).Value.ToString()

        FuncionesInfGestion.EliminaCtaEOAFDet(IdCtaEOAFDet)
        gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
        gdvCtaEOAFDet.DataBind()
        If FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value).Rows(0)("IdCtaEOAFDet").ToString = "0" Then gdvCtaEOAFDet.Rows(0).Visible = False
    End Sub

    Protected Sub gdvCtaEOAFDet_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaEOAFDet.RowCancelingEdit
        gdvCtaEOAFDet.EditIndex = -1
        gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
        gdvCtaEOAFDet.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaEOAFDet.FooterRow.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaEOAFDet.FooterRow.FindControl("cboCodCtaOrigen"), DropDownList)
        FuncionesInfGestion.AgregaCtaEOAFDet(hdfIdCtaEOAF.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue)
        gdvCtaEOAFDet.EditIndex = -1
        gdvCtaEOAFDet.DataSource = FuncionesInfGestion.LlenaEOAFDet(hdfIdCtaEOAF.Value)
        gdvCtaEOAFDet.DataBind()
    End Sub

End Class
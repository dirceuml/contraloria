Imports System.Data.OleDb
Imports System.Configuration

Public Class BG
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

        If Not IsPostBack Then
            gdvCtaBG.DataSource = LlenaBG()
            gdvCtaBG.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaBG_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaBG.RowDataBound
        Dim FlagModif As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim IdCtaBG As String = gdvCtaBG.DataKeys(e.Row.RowIndex).Value.ToString

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionEOAF
                cboCodSeccion.DataBind()
                cboCodSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()
                Dim cboFlagModif As DropDownList = CType(e.Row.FindControl("cboFlagModif"), DropDownList)
                cboFlagModif.SelectedValue = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
            Else
                FlagModif = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
                If FlagModif = "S" Then
                    FlagModif = "Sí"
                Else
                    FlagModif = "No"
                End If
                Dim lblFlagModif As Label = CType(e.Row.FindControl("lblFlagModif"), Label)
                lblFlagModif.Text = FlagModif

                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                If FlagModif = "Sí" Then
                    hlnkDetalle.NavigateUrl = "BGDetalle.aspx?IdCtaBG=" + IdCtaBG
                Else
                    hlnkDetalle.Visible = False
                End If
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
            cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionBG
            cboCodSeccion.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaBG_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaBG.RowEditing
        gdvCtaBG.EditIndex = e.NewEditIndex
        gdvCtaBG.DataSource = FuncionesInfGestion.LlenaBG()
        gdvCtaBG.DataBind()
    End Sub

    Protected Sub gdvCtaBG_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaBG.RowUpdating
        Dim IdCtaBG As String = gdvCtaBG.DataKeys(e.RowIndex).Value.ToString
        Dim row As GridViewRow = gdvCtaBG.Rows(e.RowIndex)
        Dim txtCtaBG As TextBox = DirectCast(row.FindControl("txtCtaBG"), TextBox)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)

        FuncionesInfGestion.ActualizaCtaBG(IdCtaBG, txtCtaBG.Text, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaBG.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaBG.DataBind()
    End Sub

    Protected Sub gdvCtaBG_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaBG.RowDeleting
        Dim IdCtaBG As String

        IdCtaBG = gdvCtaBG.DataKeys(e.RowIndex).Value.ToString
        FuncionesInfGestion.EliminaCtaEOAF(IdCtaBG)
        gdvCtaBG.DataSource = FuncionesInfGestion.LlenaBG()
        gdvCtaBG.DataBind()
    End Sub

    Protected Sub gdvCtaBG_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaBG.RowCancelingEdit
        gdvCtaBG.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaBG.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaBG As TextBox = DirectCast(gdvCtaBG.FooterRow.FindControl("txtCtaBG"), TextBox)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaBG.FooterRow.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaBG.FooterRow.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaBG.FooterRow.FindControl("txtOrden"), TextBox)
        FuncionesInfGestion.AgregaCtaBG(txtCtaBG.Text, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaBG.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaBG.DataBind()
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        gdvCtaBGExp.DataSource = FuncionesInfGestion.LlenaBGExp
        gdvCtaBGExp.DataBind()
        gdvCtaBGExp.Visible = True
        FuncionesVarias.DescargaExcel(Response, gdvCtaBGExp, "Ctas Balance General.xls")
    End Sub

End Class
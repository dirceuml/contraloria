Imports System.Data.OleDb
Imports System.Configuration

Public Class EOAF
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

        If Not IsPostBack Then
            gdvCtaEOAF.DataSource = LlenaEOAF()
            gdvCtaEOAF.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaEOAF_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaEOAF.RowDataBound
        Dim Signo, FlagModif As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim IdCtaEOAF As String = gdvCtaEOAF.DataKeys(e.Row.RowIndex).Value.ToString

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionEOAF
                cboCodSeccion.DataBind()
                cboCodSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()
                Dim cboFlagModif As DropDownList = CType(e.Row.FindControl("cboFlagModif"), DropDownList)
                cboFlagModif.SelectedValue = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
            Else
                Signo = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                If Signo = "1" Then
                    Signo = " + "
                ElseIf Signo = "-1" Then
                    Signo = " - "
                Else
                    Signo = "   "
                End If
                Dim lblSigno As Label = CType(e.Row.FindControl("lblSigno"), Label)
                lblSigno.Text = Signo

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
                    hlnkDetalle.NavigateUrl = "EOAFDetalle.aspx?IdCtaEOAF=" + IdCtaEOAF
                Else
                    hlnkDetalle.Visible = False
                End If
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionEOAF
            cboCodSeccion.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaEOAF_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaEOAF.RowEditing
        gdvCtaEOAF.EditIndex = e.NewEditIndex
        gdvCtaEOAF.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaEOAF.DataBind()
    End Sub

    Protected Sub gdvCtaEOAF_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaEOAF.RowUpdating
        Dim IdCtaEOAF As String = gdvCtaEOAF.DataKeys(e.RowIndex).Value.ToString
        Dim row As GridViewRow = gdvCtaEOAF.Rows(e.RowIndex)
        Dim txtCtaEOAF As TextBox = DirectCast(row.FindControl("txtCtaEOAF"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)

        FuncionesInfGestion.ActualizaCtaEOAF(IdCtaEOAF, txtCtaEOAF.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaEOAF.EditIndex = -1
        gdvCtaEOAF.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaEOAF.DataBind()
    End Sub

    Protected Sub gdvCtaEOAF_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaEOAF.RowDeleting
        Dim IdCtaEOAF As String

        IdCtaEOAF = gdvCtaEOAF.DataKeys(e.RowIndex).Value.ToString
        FuncionesInfGestion.EliminaCtaEOAF(IdCtaEOAF)
        gdvCtaEOAF.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaEOAF.DataBind()
    End Sub

    Protected Sub gdvCtaEOAF_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaEOAF.RowCancelingEdit
        gdvCtaEOAF.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaEOAF.DataBind()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaEOAF As TextBox = DirectCast(gdvCtaEOAF.FooterRow.FindControl("txtCtaEOAF"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaEOAF.FooterRow.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaEOAF.FooterRow.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaEOAF.FooterRow.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaEOAF.FooterRow.FindControl("txtOrden"), TextBox)
        FuncionesInfGestion.AgregaCtaEOAF(txtCtaEOAF.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaEOAF.DataSource = FuncionesInfGestion.LlenaEOAF()
        gdvCtaEOAF.DataBind()
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        gdvCtaEOAFExp.DataSource = FuncionesInfGestion.LlenaEOAFExp
        gdvCtaEOAFExp.DataBind()
        gdvCtaEOAFExp.Visible = True
        FuncionesVarias.DescargaExcel(Response, gdvCtaEOAFExp, "Ctas Origen Aplicacion Fondos.xls")
    End Sub

End Class
Public Class ER
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

        If Not IsPostBack Then
            gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
            gdvCtaER.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaER_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCtaER.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim IdCtaER As String = gdvCtaER.DataKeys(e.Row.RowIndex).Value.ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionER
                cboCodSeccion.DataBind()
                cboCodSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()
                Dim cboFlagModif As DropDownList = CType(e.Row.FindControl("cboFlagModif"), DropDownList)
                cboFlagModif.SelectedValue = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
            Else
                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                If DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString() = "S" Then
                    hlnkDetalle.NavigateUrl = "ERDetalle.aspx?IdCtaER=" + IdCtaER
                Else
                    hlnkDetalle.Visible = False
                End If
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
            cboCodSeccion.DataSource = FuncionesInfGestion.LlenaCodSeccionER
            cboCodSeccion.DataBind()
        End If
    End Sub

    Protected Sub gdvCtaER_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCtaER.RowEditing
        gdvCtaER.EditIndex = e.NewEditIndex
        gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
        gdvCtaER.DataBind()
    End Sub

    Protected Sub gdvCtaER_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCtaER.RowUpdating
        Dim row As GridViewRow = gdvCtaER.Rows(e.RowIndex)
        Dim IdCtaER As String = gdvCtaER.DataKeys(e.RowIndex).Value.ToString()

        Dim txtCtaER As TextBox = DirectCast(row.FindControl("txtCtaER"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)

        FuncionesInfGestion.ActualizaCtaER(IdCtaER, txtCtaER.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaER.EditIndex = -1
        gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
        gdvCtaER.DataBind()
    End Sub

    Protected Sub gdvCtaER_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvCtaER.RowDeleting
        Dim IdCtaER As String = gdvCtaER.DataKeys(e.RowIndex).Value.ToString()

        FuncionesInfGestion.EliminaCtaER(IdCtaER)
        gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
        gdvCtaER.DataBind()
    End Sub

    Protected Sub gdvCtaER_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCtaER.RowCancelingEdit
        gdvCtaER.EditIndex = -1
        gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
        gdvCtaER.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaER As TextBox = DirectCast(gdvCtaER.FooterRow.FindControl("txtCtaER"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaER.FooterRow.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaER.FooterRow.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaER.FooterRow.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaER.FooterRow.FindControl("txtOrden"), TextBox)

        FuncionesInfGestion.AgregaCtaER(txtCtaER.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaER.EditIndex = -1
        gdvCtaER.DataSource = FuncionesInfGestion.LlenaER
        gdvCtaER.DataBind()
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        gdvCtaERExp.DataSource = FuncionesInfGestion.LlenaERExp
        gdvCtaERExp.DataBind()
        gdvCtaERExp.Visible = True
        FuncionesVarias.DescargaExcel(Response, gdvCtaERExp, "Ctas Estado Resultados.xls")
    End Sub

End Class
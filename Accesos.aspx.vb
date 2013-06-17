Public Class Accesos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim CodTipo As String = ""

        If Not Session("IdUsuario") Is Nothing Then CodTipo = Session("CodTipo").ToString
        If CodTipo <> "ADM" Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
            gdvAcceso.DataBind()
        End If
    End Sub

    Protected Sub gdvAcceso_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvAcceso.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSeccion As DropDownList = CType(e.Row.FindControl("cboSeccion"), DropDownList)
                cboSeccion.DataSource = FuncionesSeguridad.LlenaSeccion()
                cboSeccion.DataBind()
                cboSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Seccion").ToString()
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboSeccion As DropDownList = CType(e.Row.FindControl("cboSeccion"), DropDownList)
            cboSeccion.DataSource = FuncionesSeguridad.LlenaSeccion()
            cboSeccion.DataBind()
        End If
    End Sub

    Protected Sub gdvAcceso_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvAcceso.RowEditing
        gdvAcceso.EditIndex = e.NewEditIndex
        gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
        gdvAcceso.DataBind()
    End Sub

    Protected Sub gdvAcceso_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvAcceso.RowUpdating
        Dim row As GridViewRow = gdvAcceso.Rows(e.RowIndex)
        Dim IdAcceso As String = gdvAcceso.DataKeys(e.RowIndex).Value.ToString()

        Dim cboSeccion As DropDownList = DirectCast(row.FindControl("cboSeccion"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        Dim txtAcceso As TextBox = DirectCast(row.FindControl("txtAcceso"), TextBox)
        Dim txtPagina As TextBox = DirectCast(row.FindControl("txtPagina"), TextBox)

        FuncionesSeguridad.ActualizaAcceso(IdAcceso, cboSeccion.SelectedValue, txtOrden.Text, txtAcceso.Text, txtPagina.Text)
        gdvAcceso.EditIndex = -1
        gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
        gdvAcceso.DataBind()
    End Sub

    Protected Sub gdvAcceso_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvAcceso.RowDeleting
        Dim IdAcceso As String = gdvAcceso.DataKeys(e.RowIndex).Value.ToString()

        FuncionesSeguridad.EliminaAcceso(IdAcceso)
        gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
        gdvAcceso.DataBind()
    End Sub

    Protected Sub gdvAcceso_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvAcceso.RowCancelingEdit
        gdvAcceso.EditIndex = -1
        gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
        gdvAcceso.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboSeccion As DropDownList = DirectCast(gdvAcceso.FooterRow.FindControl("cboSeccion"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvAcceso.FooterRow.FindControl("txtOrden"), TextBox)
        Dim txtAcceso As TextBox = DirectCast(gdvAcceso.FooterRow.FindControl("txtAcceso"), TextBox)
        Dim txtPagina As TextBox = DirectCast(gdvAcceso.FooterRow.FindControl("txtPagina"), TextBox)

        FuncionesSeguridad.AgregaAcceso(cboSeccion.SelectedValue, txtOrden.Text, txtAcceso.Text, txtPagina.Text)
        gdvAcceso.EditIndex = -1
        gdvAcceso.DataSource = FuncionesSeguridad.LlenaAccesos
        gdvAcceso.DataBind()
    End Sub

End Class
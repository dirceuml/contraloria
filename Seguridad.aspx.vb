
Public Class Seguridad
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim CodTipo As String = ""

        If Not Session("IdUsuario") Is Nothing Then CodTipo = Session("CodTipo").ToString
        If CodTipo <> "ADM" Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("IdUsuario") = "" Then Response.Redirect("Usuarios.aspx")

        If Not IsPostBack Then
            hdfIdUsuario.Value = Request.QueryString("IdUsuario")
            lblUsuario.Text = "Usuario : " & FuncionesSeguridad.BuscaUsuario(hdfIdUsuario.Value, CodTipo)
            gdvSeguridad.DataSource = FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value)
            gdvSeguridad.DataBind()
            If FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value).Rows(0)("IdSeguridad").ToString = "0" Then gdvSeguridad.Rows(0).Visible = False
        End If
    End Sub

    Protected Sub gdvSeguridad_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvSeguridad.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboAcceso As DropDownList = CType(e.Row.FindControl("cboAcceso"), DropDownList)
                cboAcceso.DataSource = LlenaAccesos()
                cboAcceso.DataBind()
                cboAcceso.SelectedValue = DataBinder.Eval(e.Row.DataItem, "SeccionAcceso").ToString()
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboAcceso As DropDownList = CType(e.Row.FindControl("cboAcceso"), DropDownList)
            cboAcceso.DataSource = LlenaAccesos()
            cboAcceso.DataBind()
        End If
    End Sub

    Protected Sub gdvSeguridad_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvSeguridad.RowEditing
        gdvSeguridad.EditIndex = e.NewEditIndex
        gdvSeguridad.DataSource = FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value)
        gdvSeguridad.DataBind()
    End Sub

    Protected Sub gdvSeguridad_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvSeguridad.RowUpdating

    End Sub

    Protected Sub gdvSeguridad_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvSeguridad.RowDeleting
        Dim IdSeguridad As String = gdvSeguridad.DataKeys(e.RowIndex).Value.ToString()

        FuncionesSeguridad.EliminaSeguridad(IdSeguridad)
        gdvSeguridad.DataSource = FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value)
        gdvSeguridad.DataBind()
        If FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value).Rows(0)("IdSeguridad").ToString = "0" Then gdvSeguridad.Rows(0).Visible = False
    End Sub

    Protected Sub gdvSeguridad_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvSeguridad.RowCancelingEdit

    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboAcceso As DropDownList = DirectCast(gdvSeguridad.FooterRow.FindControl("cboAcceso"), DropDownList)

        FuncionesSeguridad.AgregaSeguridad(hdfIdUsuario.Value, cboAcceso.SelectedValue)
        gdvSeguridad.EditIndex = -1
        gdvSeguridad.DataSource = FuncionesSeguridad.LlenaSeguridad(hdfIdUsuario.Value)
        gdvSeguridad.DataBind()
    End Sub

End Class
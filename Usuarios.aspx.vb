Public Class Usuarios
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim CodTipo As String = ""

        If Not Session("IdUsuario") Is Nothing Then CodTipo = Session("CodTipo").ToString
        If CodTipo <> "ADM" Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
            gdvUsuario.DataBind()
        End If
    End Sub

    Protected Sub gdvUsuario_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvUsuario.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim IdUsuario As String = gdvUsuario.DataKeys(e.Row.RowIndex).Value.ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboTipo As DropDownList = CType(e.Row.FindControl("cboTipo"), DropDownList)
                cboTipo.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodTipo").ToString()
            Else
                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                If DataBinder.Eval(e.Row.DataItem, "CodTipo").ToString() <> "ADM" Then
                    hlnkDetalle.NavigateUrl = "Seguridad.aspx?IdUsuario=" + IdUsuario
                Else
                    hlnkDetalle.Visible = False
                End If
            End If
        End If
    End Sub

    Protected Sub gdvUsuario_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvUsuario.RowEditing
        gdvUsuario.EditIndex = e.NewEditIndex
        gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
        gdvUsuario.DataBind()
    End Sub

    Protected Sub gdvUsuario_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvUsuario.RowUpdating
        Dim row As GridViewRow = gdvUsuario.Rows(e.RowIndex)
        Dim IdUsuario As String = gdvUsuario.DataKeys(e.RowIndex).Value.ToString()

        Dim txtUsuario As TextBox = DirectCast(row.FindControl("txtUsuario"), TextBox)
        Dim txtCodUsuario As TextBox = DirectCast(row.FindControl("txtCodUsuario"), TextBox)
        Dim txtPassword As TextBox = DirectCast(row.FindControl("txtPassword"), TextBox)
        Dim cboTipo As DropDownList = DirectCast(row.FindControl("cboTipo"), DropDownList)

        FuncionesSeguridad.ActualizaUsuario(IdUsuario, txtUsuario.Text, txtCodUsuario.Text, txtPassword.Text, cboTipo.SelectedValue)
        gdvUsuario.EditIndex = -1
        gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
        gdvUsuario.DataBind()
    End Sub

    Protected Sub gdvUsuario_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs) Handles gdvUsuario.RowDeleting
        Dim IdUsuario As String = gdvUsuario.DataKeys(e.RowIndex).Value.ToString()

        FuncionesSeguridad.EliminaUsuario(IdUsuario)
        gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
        gdvUsuario.DataBind()
    End Sub

    Protected Sub gdvUsuario_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvUsuario.RowCancelingEdit
        gdvUsuario.EditIndex = -1
        gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
        gdvUsuario.DataBind()
    End Sub

    Protected Sub lnkInsert_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtUsuario As TextBox = DirectCast(gdvUsuario.FooterRow.FindControl("txtUsuario"), TextBox)
        Dim txtCodUsuario As TextBox = DirectCast(gdvUsuario.FooterRow.FindControl("txtCodUsuario"), TextBox)
        Dim txtPassword As TextBox = DirectCast(gdvUsuario.FooterRow.FindControl("txtPassword"), TextBox)
        Dim cboTipo As DropDownList = DirectCast(gdvUsuario.FooterRow.FindControl("cboTipo"), DropDownList)

        FuncionesSeguridad.AgregaUsuario(txtUsuario.Text, txtCodUsuario.Text, txtPassword.Text, cboTipo.SelectedValue)
        gdvUsuario.EditIndex = -1
        gdvUsuario.DataSource = FuncionesSeguridad.LlenaUsuarios
        gdvUsuario.DataBind()
    End Sub

End Class
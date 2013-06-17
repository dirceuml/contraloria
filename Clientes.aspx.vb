
Public Class Clientes
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Clientes.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            hdfOrden.Value = "Cliente"
            hdfTipoOrden.Value = "asc"
            gdvClientes.PageSize = 25
            gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
            gdvClientes.DataBind()
        End If
    End Sub

    Protected Sub gdvClientes_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvClientes.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboGruposCliente As DropDownList = DirectCast(e.Row.FindControl("cboGruposCliente"), DropDownList)
                cboGruposCliente.DataSource = LlenaGruposCliente()
                cboGruposCliente.DataTextField = "GrupoCliente"
                cboGruposCliente.DataValueField = "GrupoCliente"
                cboGruposCliente.DataBind()
                cboGruposCliente.SelectedValue = DataBinder.Eval(e.Row.DataItem, "GrupoCliente").ToString()
            Else
                'nada
            End If
        End If
    End Sub

    Protected Sub gdvClientes_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvClientes.RowEditing
        gdvClientes.EditIndex = e.NewEditIndex
        gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientes.DataBind()
    End Sub

    Protected Sub gdvClientes_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvClientes.RowUpdating
        Dim IdCliente As String
        Dim GrupoCliente As String = ""
        Dim row As GridViewRow = gdvClientes.Rows(e.RowIndex)
        Dim cboGruposCliente As DropDownList = DirectCast(row.FindControl("cboGruposCliente"), DropDownList)
        Dim txtGrupoCliente As TextBox = DirectCast(row.FindControl("txtGrupoCliente"), TextBox)

        IdCliente = gdvClientes.DataKeys(e.RowIndex).Value.ToString
        If txtGrupoCliente.Text.Trim <> "" Then
            GrupoCliente = txtGrupoCliente.Text.Trim.ToUpper
        ElseIf cboGruposCliente.SelectedIndex > 0 Then
            GrupoCliente = cboGruposCliente.SelectedValue
        End If

        ActualizaGrupoCliente(IdCliente, GrupoCliente)

        gdvClientes.EditIndex = -1
        gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientes.DataBind()
    End Sub

    Protected Sub gdvClientes_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvClientes.RowCancelingEdit
        gdvClientes.EditIndex = -1
        gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientes.DataBind()
    End Sub

    Protected Sub gdvClientes_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs) Handles gdvClientes.Sorting
        If hdfOrden.Value = e.SortExpression Then
            If hdfTipoOrden.Value = "asc" Then hdfTipoOrden.Value = "desc" Else hdfTipoOrden.Value = "asc"
        Else
            hdfOrden.Value = e.SortExpression
            hdfTipoOrden.Value = "asc"
        End If
        gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientes.DataBind()
    End Sub

    Protected Sub gdvClientes_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gdvClientes.PageIndexChanging
        gdvClientes.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientes.PageIndex = e.NewPageIndex
        gdvClientes.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvClientesExp.DataSource = LlenaClientes(hdfOrden.Value, hdfTipoOrden.Value)
        gdvClientesExp.DataBind()
        gdvClientesExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvClientesExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Clientes.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
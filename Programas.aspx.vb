
Public Class Programas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Programas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        gdvProgramas.PageSize = 25
        If Not IsPostBack Then
            hdfOrden.Value = "Programa"
            hdfTipoOrden.Value = "asc"
            gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
            gdvProgramas.PageIndex = 0
            gdvProgramas.DataBind()
        End If
    End Sub

    Protected Sub gdvProgramas_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvProgramas.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboGruposPrograma As DropDownList = DirectCast(e.Row.FindControl("cboGruposPrograma"), DropDownList)
                cboGruposPrograma.DataSource = LlenaGruposPrograma()
                cboGruposPrograma.DataTextField = "GrupoPrograma"
                cboGruposPrograma.DataValueField = "GrupoPrograma"
                cboGruposPrograma.DataBind()
                cboGruposPrograma.SelectedValue = DataBinder.Eval(e.Row.DataItem, "GrupoPrograma").ToString()
            Else
                'nada
            End If
        End If
    End Sub

    Protected Sub gdvProgramas_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvProgramas.RowEditing
        gdvProgramas.EditIndex = e.NewEditIndex
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.DataBind()
    End Sub

    Protected Sub gdvProgramas_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvProgramas.RowUpdating
        Dim IdPrograma As String
        Dim GrupoPrograma As String = ""
        Dim row As GridViewRow = gdvProgramas.Rows(e.RowIndex)
        Dim cboGruposPrograma As DropDownList = DirectCast(row.FindControl("cboGruposPrograma"), DropDownList)
        Dim txtGrupoPrograma As TextBox = DirectCast(row.FindControl("txtGrupoPrograma"), TextBox)
        Dim chkFlagMaterial As CheckBox = DirectCast(row.FindControl("chkFlagMaterial"), CheckBox)

        IdPrograma = gdvProgramas.DataKeys(e.RowIndex).Value.ToString
        If txtGrupoPrograma.Text.Trim <> "" Then
            GrupoPrograma = txtGrupoPrograma.Text.Trim.ToUpper
        ElseIf cboGruposPrograma.SelectedIndex > 0 Then
            GrupoPrograma = cboGruposPrograma.SelectedValue
        End If

        ActualizaGrupoPrograma(IdPrograma, GrupoPrograma, chkFlagMaterial.Checked)

        gdvProgramas.EditIndex = -1
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.DataBind()
    End Sub

    Protected Sub gdvProgramas_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvProgramas.RowCancelingEdit
        gdvProgramas.EditIndex = -1
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.DataBind()
    End Sub

    Protected Sub gdvProgramas_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs) Handles gdvProgramas.Sorting
        If hdfOrden.Value = e.SortExpression Then
            If hdfTipoOrden.Value = "asc" Then hdfTipoOrden.Value = "desc" Else hdfTipoOrden.Value = "asc"
        Else
            hdfOrden.Value = e.SortExpression
            hdfTipoOrden.Value = "asc"
        End If
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.PageIndex = 0
        gdvProgramas.DataBind()
    End Sub

    Protected Sub gdvProgramas_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gdvProgramas.PageIndexChanging
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.PageIndex = e.NewPageIndex
        gdvProgramas.DataBind()
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click
        hdfOrden.Value = "Programa"
        hdfTipoOrden.Value = "asc"
        gdvProgramas.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramas.PageIndex = 0
        gdvProgramas.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvProgramasExp.DataSource = LlenaProgramas(txtProgramaB.Text, hdfOrden.Value, hdfTipoOrden.Value)
        gdvProgramasExp.DataBind()
        gdvProgramasExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvProgramasExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Programas.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub
End Class
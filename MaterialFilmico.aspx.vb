
Public Class MaterialFilmico
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "MaterialFilmico.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        gdvMaterialFilmico.PageSize = 25
        If Not IsPostBack Then
            hdfOrden.Value = "Material"
            hdfTipoOrden.Value = "asc"
            cboNumContrato.DataSource = LlenaNumContrato()
            cboNumContrato.DataTextField = "NumContrato2"
            cboNumContrato.DataValueField = "NumContrato"
            cboNumContrato.DataBind()
            gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
            gdvMaterialFilmico.PageIndex = 0
            gdvMaterialFilmico.DataBind()
        End If
    End Sub

    Protected Sub gdvMaterialFilmico_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvMaterialFilmico.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboGruposPrograma As DropDownList = DirectCast(e.Row.FindControl("cboGruposPrograma"), DropDownList)
                cboGruposPrograma.DataSource = LlenaGruposProgramaMaterial()
                cboGruposPrograma.DataTextField = "GrupoPrograma"
                cboGruposPrograma.DataValueField = "GrupoPrograma"
                cboGruposPrograma.DataBind()
                cboGruposPrograma.SelectedValue = DataBinder.Eval(e.Row.DataItem, "GrupoPrograma").ToString()
            Else
                Dim CodMaterial As String
                CodMaterial = DataBinder.Eval(e.Row.DataItem, "CodMaterial").ToString()
                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                hlnkDetalle.NavigateUrl = "ProgMaterialFilmico.aspx?CodMaterial=" + CodMaterial
            End If
        End If
    End Sub

    Protected Sub gdvMaterialFilmico_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvMaterialFilmico.RowEditing
        gdvMaterialFilmico.EditIndex = e.NewEditIndex
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub gdvMaterialFilmico_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvMaterialFilmico.RowUpdating
        Dim IdMaterialFilmico As String
        Dim GrupoPrograma As String = ""
        Dim row As GridViewRow = gdvMaterialFilmico.Rows(e.RowIndex)
        Dim cboGruposPrograma As DropDownList = DirectCast(row.FindControl("cboGruposPrograma"), DropDownList)
        Dim txtGrupoPrograma As TextBox = DirectCast(row.FindControl("txtGrupoPrograma"), TextBox)

        IdMaterialFilmico = gdvMaterialFilmico.DataKeys(e.RowIndex).Value.ToString
        If txtGrupoPrograma.Text.Trim <> "" Then
            GrupoPrograma = txtGrupoPrograma.Text.Trim.ToUpper
        ElseIf cboGruposPrograma.SelectedIndex > 0 Then
            GrupoPrograma = cboGruposPrograma.SelectedValue
        End If

        ActualizaGrupoPrograma(IdMaterialFilmico, GrupoPrograma)

        gdvMaterialFilmico.EditIndex = -1
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub gdvMaterialFilmico_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvMaterialFilmico.RowCancelingEdit
        gdvMaterialFilmico.EditIndex = -1
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub gdvMaterialFilmico_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs) Handles gdvMaterialFilmico.Sorting
        If hdfOrden.Value = e.SortExpression Then
            If hdfTipoOrden.Value = "asc" Then hdfTipoOrden.Value = "desc" Else hdfTipoOrden.Value = "asc"
        Else
            hdfOrden.Value = e.SortExpression
            hdfTipoOrden.Value = "asc"
        End If
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.PageIndex = 0
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub gdvMaterialFilmico_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gdvMaterialFilmico.PageIndexChanging
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.PageIndex = e.NewPageIndex
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click
        hdfOrden.Value = "Material"
        hdfTipoOrden.Value = "asc"
        gdvMaterialFilmico.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmico.PageIndex = 0
        gdvMaterialFilmico.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvMaterialFilmicoExp.DataSource = LlenaMaterialFilmico(txtMaterialB.Text, cboNumContrato.SelectedValue, hdfOrden.Value, hdfTipoOrden.Value)
        gdvMaterialFilmicoExp.DataBind()
        gdvMaterialFilmicoExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvMaterialFilmicoExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Material Filmico.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
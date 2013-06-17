
Public Class CentrosCosto
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "CentrosCosto.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            hdfOrden.Value = "CodCentroCosto"
            hdfTipoOrden.Value = "asc"
            gdvCentrosCosto.PageSize = 25
            gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
            gdvCentrosCosto.DataBind()
        End If
    End Sub

    'Protected Sub gdvCentroCosto_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvCentrosCosto.RowDataBound
    '    If e.Row.RowType = DataControlRowType.DataRow Then
    '        If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
    '            'nada
    '        Else
    '            'nada
    '        End If
    '    End If
    'End Sub

    Protected Sub gdvCentroCosto_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvCentrosCosto.RowEditing
        gdvCentrosCosto.EditIndex = e.NewEditIndex
        gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCosto.DataBind()
    End Sub

    Protected Sub gdvCentroCosto_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvCentrosCosto.RowUpdating
        Dim CodCentroCosto As String
        Dim row As GridViewRow = gdvCentrosCosto.Rows(e.RowIndex)
        Dim cboTipo As DropDownList = DirectCast(row.FindControl("cboTipo"), DropDownList)
        Dim chkFlagDeporte As CheckBox = DirectCast(row.FindControl("chkFlagDeporte"), CheckBox)
        Dim txtCentroCostoEGP As TextBox = DirectCast(row.FindControl("txtCentroCostoEGP"), TextBox)
        Dim txtCodGrupoEGP As TextBox = DirectCast(row.FindControl("txtCodGrupoEGP"), TextBox)
        Dim txtGrupoEGP As TextBox = DirectCast(row.FindControl("txtGrupoEGP"), TextBox)

        CodCentroCosto = gdvCentrosCosto.DataKeys(e.RowIndex).Value.ToString

        ActualizaCentroCosto(CodCentroCosto, cboTipo.SelectedValue, chkFlagDeporte.Checked, txtCentroCostoEGP.Text, txtCodGrupoEGP.Text, txtGrupoEGP.Text)

        gdvCentrosCosto.EditIndex = -1
        gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCosto.DataBind()
    End Sub

    Protected Sub gdvCentroCosto_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvCentrosCosto.RowCancelingEdit
        gdvCentrosCosto.EditIndex = -1
        gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCosto.DataBind()
    End Sub

    Protected Sub gdvCentroCosto_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs) Handles gdvCentrosCosto.Sorting
        If hdfOrden.Value = e.SortExpression Then
            If hdfTipoOrden.Value = "asc" Then hdfTipoOrden.Value = "desc" Else hdfTipoOrden.Value = "asc"
        Else
            hdfOrden.Value = e.SortExpression
            hdfTipoOrden.Value = "asc"
        End If
        gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCosto.PageIndex = 0
        gdvCentrosCosto.DataBind()
    End Sub

    Protected Sub gdvCentroCosto_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gdvCentrosCosto.PageIndexChanging
        gdvCentrosCosto.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCosto.PageIndex = e.NewPageIndex
        gdvCentrosCosto.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvCentrosCostoExp.DataSource = LlenaCentrosCosto(hdfOrden.Value, hdfTipoOrden.Value)
        gdvCentrosCostoExp.DataBind()
        gdvCentrosCostoExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvCentrosCostoExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Centros Costo.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class

Public Class Pasivos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Pasivos.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvPasivos.DataSource = LlenaPasivos()
            gdvPasivos.DataBind()
        End If
    End Sub

    Protected Sub lnkNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboGrupo As DropDownList = DirectCast(gdvPasivos.FooterRow.FindControl("cboGrupo"), DropDownList)
        Dim cboSubGrupo As DropDownList = DirectCast(gdvPasivos.FooterRow.FindControl("cboSubGrupo"), DropDownList)
        Dim cboCuenta As DropDownList = DirectCast(gdvPasivos.FooterRow.FindControl("cboCuenta"), DropDownList)
        Dim cboSeccion As DropDownList = DirectCast(gdvPasivos.FooterRow.FindControl("cboSeccion"), DropDownList)
        Dim chkFlagGpoATV As CheckBox = DirectCast(gdvPasivos.FooterRow.FindControl("chkFlagGpoATV"), CheckBox)

        FuncionesPasivos.CreaPasivo(cboSeccion.SelectedValue, cboSeccion.SelectedItem.Text, cboGrupo.SelectedValue, cboGrupo.SelectedItem.Text, _
                                    cboSubGrupo.SelectedValue, cboSubGrupo.SelectedItem.Text, cboCuenta.SelectedValue, chkFlagGpoATV.Checked)

        gdvPasivos.EditIndex = -1
        gdvPasivos.DataSource = LlenaPasivos()
        gdvPasivos.DataBind()
    End Sub

    Protected Sub gdvPasivos_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvPasivos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSeccion As DropDownList = DirectCast(e.Row.FindControl("cboSeccion"), DropDownList)
                Dim cboGrupo As DropDownList = DirectCast(e.Row.FindControl("cboGrupo"), DropDownList)
                Dim cboSubGrupo As DropDownList = DirectCast(e.Row.FindControl("cboSubGrupo"), DropDownList)
                Dim cboCuenta As DropDownList = DirectCast(e.Row.FindControl("cboCuenta"), DropDownList)

                cboSeccion.DataSource = FuncionesPasivos.LlenaSeccion()
                cboSeccion.DataTextField = "Seccion"
                cboSeccion.DataValueField = "CodSeccion"
                cboSeccion.DataBind()
                cboSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()

                cboGrupo.DataSource = LlenaGrupo()
                cboGrupo.DataTextField = "Grupo"
                cboGrupo.DataValueField = "CodGrupo"
                cboGrupo.DataBind()
                cboGrupo.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodGrupo").ToString()

                cboSubGrupo.DataSource = LlenaSubGrupo()
                cboSubGrupo.DataTextField = "SubGrupo"
                cboSubGrupo.DataValueField = "CodSubGrupo"
                cboSubGrupo.DataBind()
                cboSubGrupo.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSubGrupo").ToString()

                cboCuenta.DataSource = LlenaCuentas_Pasivos()
                cboCuenta.DataTextField = "Cuenta"
                cboCuenta.DataValueField = "CodCuenta"
                cboCuenta.DataBind()
                cboCuenta.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCuenta").ToString()
            Else
                'nada
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboSeccion As DropDownList = DirectCast(e.Row.FindControl("cboSeccion"), DropDownList)
            Dim cboGrupo As DropDownList = DirectCast(e.Row.FindControl("cboGrupo"), DropDownList)
            Dim cboSubGrupo As DropDownList = DirectCast(e.Row.FindControl("cboSubGrupo"), DropDownList)
            Dim cboCuenta As DropDownList = DirectCast(e.Row.FindControl("cboCuenta"), DropDownList)

            cboSeccion.DataSource = FuncionesPasivos.LlenaSeccion()
            cboSeccion.DataTextField = "Seccion"
            cboSeccion.DataValueField = "CodSeccion"
            cboSeccion.DataBind()

            cboGrupo.DataSource = LlenaGrupo()
            cboGrupo.DataTextField = "Grupo"
            cboGrupo.DataValueField = "CodGrupo"
            cboGrupo.DataBind()

            cboSubGrupo.DataSource = LlenaSubGrupo()
            cboSubGrupo.DataTextField = "SubGrupo"
            cboSubGrupo.DataValueField = "CodSubGrupo"
            cboSubGrupo.DataBind()

            cboCuenta.DataSource = LlenaCuentas_Pasivos()
            cboCuenta.DataTextField = "Cuenta"
            cboCuenta.DataValueField = "CodCuenta"
            cboCuenta.DataBind()
        End If
    End Sub

    Protected Sub gdvPasivos_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs) Handles gdvPasivos.RowEditing
        gdvPasivos.EditIndex = e.NewEditIndex
        gdvPasivos.DataSource = LlenaPasivos()
        gdvPasivos.DataBind()
    End Sub

    Protected Sub gdvPasivos_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs) Handles gdvPasivos.RowUpdating
        Dim IdGrupoPasivos As String
        Dim row As GridViewRow = gdvPasivos.Rows(e.RowIndex)
        Dim cboGrupo As DropDownList = DirectCast(row.FindControl("cboGrupo"), DropDownList)
        Dim cboSubGrupo As DropDownList = DirectCast(row.FindControl("cboSubGrupo"), DropDownList)
        Dim cboCuenta As DropDownList = DirectCast(row.FindControl("cboCuenta"), DropDownList)
        Dim cboSeccion As DropDownList = DirectCast(row.FindControl("cboSeccion"), DropDownList)
        Dim chkFlagGpoATV As CheckBox = DirectCast(row.FindControl("chkFlagGpoATV"), CheckBox)

        IdGrupoPasivos = gdvPasivos.DataKeys(e.RowIndex).Value.ToString

        FuncionesPasivos.ActualizaPasivo(IdGrupoPasivos, cboSeccion.SelectedValue, cboSeccion.SelectedItem.Text, cboGrupo.SelectedValue, cboGrupo.SelectedItem.Text, _
                                    cboSubGrupo.SelectedValue, cboSubGrupo.SelectedItem.Text, cboCuenta.SelectedValue, chkFlagGpoATV.Checked)

        gdvPasivos.EditIndex = -1
        gdvPasivos.DataSource = LlenaPasivos()
        gdvPasivos.DataBind()
    End Sub

    Protected Sub gdvPasivos_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs) Handles gdvPasivos.RowCancelingEdit
        gdvPasivos.EditIndex = -1
        gdvPasivos.DataSource = LlenaPasivos()
        gdvPasivos.DataBind()
    End Sub

    Protected Sub gdvPasivos_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gdvPasivos.RowDeleting
        Dim IdGrupoPasivos As String
        Dim row As GridViewRow = gdvPasivos.Rows(e.RowIndex)

        IdGrupoPasivos = gdvPasivos.DataKeys(e.RowIndex).Value.ToString

        FuncionesPasivos.EliminaPasivo(IdGrupoPasivos)

        gdvPasivos.EditIndex = -1
        gdvPasivos.DataSource = LlenaPasivos()
        gdvPasivos.DataBind()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvPasivosExp.DataSource = LlenaPasivos()
        gdvPasivosExp.DataBind()
        gdvPasivosExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvPasivosExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Pasivos.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
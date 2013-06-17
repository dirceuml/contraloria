
Public Class ProgMaterialFilmico
    Inherits System.Web.UI.Page

    Dim CodMaterial As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "MaterialFilmico.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("CodMaterial") = "" Then Response.Redirect("MaterialFilmico.aspx")

        CodMaterial = Request.QueryString("CodMaterial")
        If Not IsPostBack Then
            lblMaterial.Text = BuscaMaterialFilmico(CodMaterial)
            gdvConsumoMaterialFilmico.DataSource = LlenaConsumoMaterialFilmico(CodMaterial)
            gdvConsumoMaterialFilmico.DataBind()
            gdvProgMaterialFilmico.DataSource = LlenaProgMaterialFilmico(CodMaterial)
            gdvProgMaterialFilmico.DataBind()
        End If
    End Sub

    Protected Sub gdvProgMaterialFilmico_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles gdvProgMaterialFilmico.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                'nada
            Else
                'nada
                Dim FechaProg As String
                FechaProg = Convert.ToDateTime(DataBinder.Eval(e.Row.DataItem, "FechaProg")).ToString("dddd dd/MM/yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-PE"))
                e.Row.Cells(1).Text = FechaProg
            End If
        End If
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExcel.Click
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        gdvConsumoMaterialFilmicoExp.DataSource = LlenaConsumoMaterialFilmico(CodMaterial)
        gdvConsumoMaterialFilmicoExp.DataBind()
        gdvConsumoMaterialFilmicoExp.Parent.Controls.Add(frm)

        gdvProgMaterialFilmicoExp.DataSource = LlenaProgMaterialFilmico(CodMaterial)
        gdvProgMaterialFilmicoExp.DataBind()
        gdvProgMaterialFilmicoExp.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvConsumoMaterialFilmicoExp)
        frm.Controls.Add(gdvProgMaterialFilmicoExp)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=Consumo Material Filmico.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

End Class
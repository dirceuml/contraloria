Public Class ReportesCtasCobrar
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesCtasCobrar.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
            txtFecha.Text = Convert.ToDateTime(Date.Now.ToString("yyyy-MM") & "-01").AddDays(-1).ToString("yyyy-MM-dd")
        End If
    End Sub

    Protected Sub lnkGenerarReporteCtasPorCobrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteCtasPorCobrar.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "CtasPorCobrar.xls"
        NombreArchivo = "Ctas por Cobrar Contabilidad " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesCtasPorCobrar(RutaPlantilla, RutaArchivo, IdPeriodo, False)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteCtasPorCobrarContraloria_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteCtasPorCobrarContraloria.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "CtasPorCobrar.xls"
        NombreArchivo = "Ctas por Cobrar Contraloria " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesCtasPorCobrar(RutaPlantilla, RutaArchivo, IdPeriodo, True)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteFacturasPorCobrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteFacturasPorCobrar.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        'Dim IdPeriodo As String

        'IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FacturasPorCobrar.xls"
        NombreArchivo = "Facturas por Cobrar " & txtFecha.Text & ".xls"
        'NombreArchivo = "Facturas por Cobrar " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFacturasPorCobrar(RutaPlantilla, RutaArchivo, txtFecha.Text)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

End Class
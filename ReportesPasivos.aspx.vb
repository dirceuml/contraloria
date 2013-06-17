Public Class ReportesPasivos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesPasivos.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
        End If
    End Sub

    Protected Sub lnkGenerarReportePasivos_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReportePasivos.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Pasivos.xls"
        NombreArchivo = "Pasivos " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesPasivos(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

End Class
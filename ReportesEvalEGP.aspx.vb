Public Class ReportesEvalEGP
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesEvalEGP.aspx"
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

    Protected Sub lnkGenerarReportePptoEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReportePptoEGP.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdAño As String

        IdAño = cboPeriodo.SelectedValue.Substring(0, 4)
        NombrePlantilla = "Presupuesto EGP.xls"
        NombreArchivo = "Presupuesto EGP " & IdAño & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesEvalEGP.GeneraReportesEGPPpto(RutaPlantilla, RutaArchivo, IdAño)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteEvalEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteEvalEGP.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Evaluacion EGP.xls"
        NombreArchivo = "Evaluacion EGP " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesEvalEGP.GeneraReportesEGPProy(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

End Class
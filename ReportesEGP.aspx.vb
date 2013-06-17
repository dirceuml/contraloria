Public Class ReportesEGP
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesEGP.aspx"
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

    Protected Sub lnkGenerarReporte_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporte.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "EGP.xls"
        NombreArchivo = "EGP " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesEGP(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteComparativoCuentas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteComparativoCuentas.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo, CodCentroCosto, IdCtaEGP As String

        IdPeriodo = cboPeriodo.SelectedValue
        CodCentroCosto = cboCentroCosto.SelectedValue
        IdCtaEGP = cboRubroCosto.SelectedValue
        NombrePlantilla = "EGP Comparativo Cuentas.xls"
        NombreArchivo = "EGP Comparativo Cuentas " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & " " & cboCentroCosto.SelectedItem.Text & " - " & cboRubroCosto.SelectedItem.Text & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesComparativos(RutaPlantilla, RutaArchivo, IdPeriodo, CodCentroCosto, IdCtaEGP)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
        cboCentroCosto.DataSource = CargaCentrosCosto()
        cboCentroCosto.DataTextField = "CentroCostoEGP"
        cboCentroCosto.DataValueField = "CodCentroCosto"
        cboCentroCosto.DataBind()
        cboRubroCosto.DataSource = CargaRubrosCosto()
        cboRubroCosto.DataTextField = "CtaEGP"
        cboRubroCosto.DataValueField = "IdCtaEGP"
        cboRubroCosto.DataBind()
    End Sub

End Class
Public Class ReportesEvolVentas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesEvolVentas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            txtFecha.Text = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
        End If
    End Sub

    Protected Sub lnkGenerarReporteEvolucionVentas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteEvolucionVentas.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim Fecha As Date

        Fecha = Convert.ToDateTime(txtFecha.Text)
        NombrePlantilla = "Evolucion Ventas.xls"
        If Fecha.ToString("yyyyMM") <> Fecha.AddDays(1).ToString("yyyyMM") Then
            NombreArchivo = "Evolucion Ventas " & Fecha.ToString("yyyy-MM") & ".xls"
        Else
            NombreArchivo = "Evolucion Ventas al " & Fecha.ToString("yyyy-MM-dd") & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesEvolucionVentas(RutaPlantilla, RutaArchivo, Fecha)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

End Class
Public Class ReportesRentabilidad
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesRentabilidad.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
            txtPeso.Text = "100"
            txtPeso2.Text = "0"
        End If
    End Sub

    Protected Sub lnkGenerarReporteRentabilidad_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteRentabilidad.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Rentabilidad.xls"
        NombreArchivo = "Rentabilidad " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesRentabilidad(RutaPlantilla, RutaArchivo, IdPeriodo, Convert.ToDouble(txtPeso.Text))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

    Protected Sub txtPeso_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtPeso.TextChanged
        If txtPeso.Text = "" Then txtPeso.Text = 0
        txtPeso2.Text = (100 - Convert.ToInt32(txtPeso.Text)).ToString
    End Sub
End Class
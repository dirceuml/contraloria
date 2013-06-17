
Public Class ImportacionEvalEGP
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ImportacionEvalEGP.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboAño.SelectedValue = Date.Now.ToString("yyyy")
        End If
    End Sub

    Sub CargaCombos()
        cboAño.DataSource = LlenaAños()
        cboAño.DataTextField = "IdAño"
        cboAño.DataValueField = "IdAño"
        cboAño.DataBind()
    End Sub

    Protected Sub lnkDescargaEvalEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDescargaEvalEGP.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String

        NombrePlantilla = "Evaluacion EGP Datos.xls"
        NombreArchivo = "Evaluacion EGP Datos " & cboAño.SelectedValue & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesEvalEGP.GeneraPlantillaEGPPpto(RutaPlantilla, RutaArchivo, Convert.ToInt32(cboAño.SelectedValue))
        FuncionesVarias.GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkImportaEvalEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportaEvalEGP.Click
        Dim flag As Boolean
        Dim log As String = ""

        fupEvalEGP.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\Evaluacion EGP Datos.xls")

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = FuncionesVarias.EjecutaProcesoCarga("EvaluacionEGP.dtsx", log)

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & "<br>" & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

End Class
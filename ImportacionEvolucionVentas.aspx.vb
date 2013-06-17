'Imports Microsoft.SqlServer.Dts.Runtime

Public Class ImportacionEvolucionVentas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ImportacionEvolucionVentas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
            cboAño.SelectedValue = Date.Now.ToString("yyyy")
        End If
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()

        cboAño.DataSource = LlenaAños()
        cboAño.DataTextField = "IdAño"
        cboAño.DataValueField = "IdAño"
        cboAño.DataBind()
    End Sub

    Protected Sub btnImportaEvolucionVentas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaEvolucionVentas.Click
        Dim flag As Boolean
        Dim log As String = ""
        Dim IdPeriodo, Fecha As String

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        IdPeriodo = cboPeriodo.SelectedValue
        'IdPeriodoAnt = (Convert.ToInt32(IdPeriodo.Substring(0, 4)) - 1).ToString & IdPeriodo.Substring(4, 2)
        Fecha = Convert.ToDateTime(IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
        'FechaAnt = Convert.ToDateTime(IdPeriodoAnt.Substring(0, 4) & "-" & IdPeriodoAnt.Substring(4, 2) & "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")

        flag = EjecutaProcesoCarga("Ventas.dtsx", log, IdPeriodo.Substring(0, 4) & "-01-01", Fecha)
        flag = flag And EjecutaProcesoCarga("LibroInventario.dtsx", log, IdPeriodo, Fecha)

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & "<br>" & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub lnkDescargaVentaPpto_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDescargaVentaPpto.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Venta Ppto.xls"
        NombreArchivo = "Venta Ppto " & IdPeriodo.Substring(0, 4) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesEvolucionVentas.GeneraPlantillaVentaPpto(RutaPlantilla, RutaArchivo, Convert.ToInt32(IdPeriodo.Substring(0, 4)))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkImportaVentaPpto_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportaVentaPpto.Click
        Dim flag As Boolean
        Dim log As String = ""

        fupVentaPpto.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\Venta Ppto.xls")

        lblEstadoPpto.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = EjecutaProcesoCarga("VentaPpto.dtsx", log)

        If flag Then
            lblEstadoPpto.Text = lblEstadoPpto.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstadoPpto.Text = lblEstadoPpto.Text & "<br>" & log
            lblEstadoPpto.Text = lblEstadoPpto.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub lnkDescargaVenta2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDescargaVenta2.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String

        NombrePlantilla = "Venta2.xls"
        NombreArchivo = "Ventas Global y ATV_SUR.xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesEvolucionVentas.GeneraPlantillaVenta2(RutaPlantilla, RutaArchivo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkImportaVenta2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportaVenta2.Click
        Dim flag As Boolean
        Dim log As String = ""

        fupVenta2.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\Venta2.xls")

        lblEstado2.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = EjecutaProcesoCarga("Venta2.dtsx", log)

        If flag Then
            lblEstado2.Text = lblEstado2.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado2.Text = lblEstado2.Text & "<br>" & log
            lblEstado2.Text = lblEstado2.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

End Class
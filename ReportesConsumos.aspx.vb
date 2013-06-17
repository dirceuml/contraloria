
Public Class ReportesConsumos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesConsumos.aspx"
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

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

    Protected Sub lnkImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportar.Click
        Dim flag As Boolean = True
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        fupConsumos.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\Consumos.xls")

        flag = EjecutaProcesoCarga("Consumos.dtsx", log)

        If flag Then
            FuncionesConsumos.ProcesoCargaConsumos(cboPeriodo.SelectedValue)
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub lnkImportarInactivos_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportarInactivos.Click
        Dim flag As Boolean = True
        Dim log As String = ""

        lblEstadoInactivos.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        fupConsumosInactivos.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\ConsumosClientesInactivos.xls")

        FuncionesConsumos.EliminaInactivos(cboPeriodo.SelectedValue)

        flag = FuncionesVarias.EjecutaProcesoCarga("ConsumosClientesInactivos.dtsx", log)

        If flag Then
            FuncionesConsumos.ProcesoCargaConsumos(cboPeriodo.SelectedValue)
            lblEstadoInactivos.Text = lblEstadoInactivos.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstadoInactivos.Text = lblEstadoInactivos.Text & log
            lblEstadoInactivos.Text = lblEstadoInactivos.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub lnkEjemploInactivos_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkEjemploInactivos.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue

        NombrePlantilla = "Consumos Clientes Inactivos.xls"
        NombreArchivo = "Consumos Clientes Inactivos " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesConsumos.GeneraPlantillaConsumosInactivos(RutaPlantilla, RutaArchivo, IdPeriodo)
        FuncionesVarias.GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporte_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporte.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Consumos.xls"
        NombreArchivo = "Consumos " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesConsumos(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    'Function EjecutarProcesosCarga(ByVal package_name As String)

    '    Dim app As Application
    '    Dim package As Package
    '    Dim results As DTSExecResult
    '    Dim path As String
    '    Dim flag As Boolean

    '    flag = True
    '    path = "E:\ATV\ATVContraloriaIS2\bin\"
    '    Try
    '        app = New Application
    '        package = app.LoadPackage(path & package_name, Nothing)
    '        results = package.Execute()
    '        If results = DTSExecResult.Failure Then
    '            Dim dtserror As DtsError
    '            For Each dtserror In package.Errors
    '                'lblEstado.Text = lblEstado.Text & package_name & ": " & dtserror.Description & "<br>"
    '            Next
    '            flag = False
    '        End If
    '    Catch ex As Exception
    '        'lblEstado.Text = lblEstado.Text & package_name & ": " & ex.Message & "<br>"
    '        flag = False
    '    End Try

    '    Return flag
    'End Function

End Class
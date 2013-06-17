'Imports Microsoft.SqlServer.Dts.Runtime

Public Class ImportacionVentas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ImportacionVentas.aspx"
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
        cboPeriodo.DataSource = LlenaPeriodosVentas()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

    Protected Sub btnImportaVentas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaVentas.Click
        Dim flag As Boolean
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = FuncionesVarias.EjecutaProcesoCarga("Facturacion.dtsx", log, cboPeriodo.SelectedValue)

        FuncionesVentas.ProcesaFacturacion(Convert.ToInt32(cboPeriodo.SelectedValue))

        flag = FuncionesVarias.EjecutaProcesoCarga("MaterialFilmico.dtsx", log)

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & "<br>" & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub lnkDescargaRating_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDescargaRating.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        FuncionesRentabilidad.DistribucionCostos(0)

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "Rating.xls"
        NombreArchivo = "Plantilla Rating " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        FuncionesVentas.GeneraPlantillaRating(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkImportaRating_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportaRating.Click
        Dim flag As Boolean
        Dim log As String = ""

        fupRating.SaveAs("E:\ATV\ATVContraloriaWeb\Temp\Rating.xls")

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = FuncionesVarias.EjecutaProcesoCarga("Rating.dtsx", log, cboPeriodo.SelectedValue)

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & "<br>" & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    'Function EjecutarProcesosCarga(ByVal package_name As String)
    '    Dim app As Application
    '    Dim package As Package
    '    Dim results As DTSExecResult
    '    Dim path As String
    '    Dim flag As Boolean

    '    flag = True
    '    path = "E:\ATV\ATVContraloriaIS\bin\"

    '    Try
    '        app = New Application
    '        package = app.LoadPackage(path & package_name, Nothing)
    '        If package_name = "Facturacion.dtsx" Then
    '            package.Variables("IdPeriodoCarga").Value = cboPeriodo.SelectedValue
    '        Else
    '            package.Variables("IdPeriodo").Value = cboPeriodo.SelectedValue
    '        End If
    '        results = package.Execute()
    '        If results = DTSExecResult.Failure Then
    '            Dim dtserror As DtsError
    '            For Each dtserror In package.Errors
    '                lblEstado.Text = lblEstado.Text & package_name & ": " & dtserror.Description & "<br>"
    '            Next
    '            flag = False
    '        End If
    '    Catch ex As Exception
    '        lblEstado.Text = lblEstado.Text & package_name & ": " & ex.Message & "<br>"
    '        flag = False
    '    End Try

    '    Return flag
    'End Function

End Class
'Imports Microsoft.SqlServer.Dts.Runtime

Public Class ImportacionEGP
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ImportacionEGP.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")
    End Sub

    Protected Sub btnImportaEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaEGP.Click
        Dim flag As Boolean
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = FuncionesVarias.EjecutaProcesoCarga("Asientos.dtsx", log)
        flag = FuncionesVarias.EjecutaProcesoCarga("FacturasPorCobrar.dtsx", log)
        FuncionesEGP.ProcesoCargaEGP()
        FuncionesCtasPorCobrar.CreaFacturaPorCobrar(Date.Today.ToString("yyyy-MM-dd"))

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
    '    If package_name <> "CierreCaja.dtsx" And package_name <> "Asientos.dtsx" Then
    '        path = "E:\ATV\ATVContraloriaIS\bin\"
    '    Else
    '        path = "E:\ATV\ATVContraloriaIS2\bin\"
    '    End If
    '    Try
    '        app = New Application
    '        package = app.LoadPackage(path & package_name, Nothing)
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
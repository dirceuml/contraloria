'Imports Microsoft.SqlServer.Dts.Runtime

Public Class ImportacionFC
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ImportacionFC.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")
    End Sub

    Protected Sub btnImportaFC_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaFC.Click
        Dim flag As Boolean = True
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = EjecutaProcesoCarga("CentrosCosto.dtsx", log)
        flag = flag And EjecutaProcesoCarga("Movimientos.dtsx", log)
        flag = flag And EjecutaProcesoCarga("LibroMayor.dtsx", log)
        flag = flag And EjecutaProcesoCarga("CierreCaja.dtsx", log)

        Dim Fecha As Date
        Fecha = Convert.ToDateTime("2013-01-01")
        While Fecha <= Date.Now.AddDays(-1)
            CalculaAjustesTipoCambio(Fecha)
            Fecha = Fecha.AddDays(1)
        End While

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub btnImportaFC2012_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaFC2012.Click
        Dim flag As Boolean = True
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = flag And EjecutaProcesoCarga("Movimientos2012.dtsx", log)
        flag = flag And EjecutaProcesoCarga("CierreCaja.dtsx", log)

        Dim Fecha As Date
        Fecha = Convert.ToDateTime("2012-01-01")
        While Fecha <= Convert.ToDateTime("2012-12-31")
            CalculaAjustesTipoCambio(Fecha)
            Fecha = Fecha.AddDays(1)
        End While

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub btnImportaFC2011_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaFC2011.Click
        Dim flag As Boolean
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = EjecutaProcesoCarga("Movimientos2011.dtsx", log)
        flag = flag And EjecutaProcesoCarga("CierreCaja.dtsx", log)

        Dim Fecha As Date
        Fecha = Convert.ToDateTime("2011-08-01")
        While Fecha <= Convert.ToDateTime("2011-12-31")
            CalculaAjustesTipoCambio(Fecha)
            Fecha = Fecha.AddDays(1)
        End While

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    Protected Sub btnImportaFC2010_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaFC2010.Click
        Dim flag As Boolean
        Dim log As String = ""

        lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

        flag = EjecutaProcesoCarga("Movimientos2010.dtsx", log)
        flag = flag And EjecutaProcesoCarga("CierreCaja.dtsx", log)

        Dim Fecha As Date
        Fecha = Convert.ToDateTime("2010-01-01")
        While Fecha <= Convert.ToDateTime("2010-12-31")
            CalculaAjustesTipoCambio(Fecha)
            Fecha = Fecha.AddDays(1)
        End While

        If flag Then
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        Else
            lblEstado.Text = lblEstado.Text & log
            lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
        End If
    End Sub

    'Protected Sub btnImportaEGP_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnImportaEGP.Click
    '    Dim flag As Boolean

    '    lblEstado.Text = "Proceso de carga iniciado (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"

    '    flag = EjecutarProcesosCarga("Asientos.dtsx")
    '    flag = EjecutarProcesosCarga("FacturasPorCobrar.dtsx")
    '    FuncionesEGPRep.ProcesoCargaEGP()

    '    If flag Then
    '        lblEstado.Text = lblEstado.Text & "<br>Proceso de carga finalizado satisfactoriamente (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
    '    Else
    '        lblEstado.Text = lblEstado.Text & "<br>Proceso de carga con errores (" & Date.Now.ToString("dd/MM/yyyy hh:mm tt") & ")"
    '    End If
    'End Sub

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
Public Class ReportesFC
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ReportesFC.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            cboPeriodo.SelectedValue = Date.Now.AddMonths(-1).ToString("yyyyMM")
            txtFechaIni.Text = Date.Now.ToString("yyyy-MM") & "-01"
            If Date.Now.DayOfWeek = DayOfWeek.Monday Then txtFechaFin.Text = Date.Now.AddDays(-3).ToString("yyyy-MM-dd") Else txtFechaFin.Text = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
        End If
    End Sub

    Protected Sub lnkGenerarReporte_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporte.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCaja.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Contabilidad " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        Else
            NombreArchivo = "Flujo Caja Contabilidad Exterior " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFC(True, rdbAmbito.SelectedValue = "NAC", RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteTesoreria_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteTesoreria.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCaja.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Tesoreria " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        Else
            NombreArchivo = "Flujo Caja Tesoreria Exterior " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFC(False, rdbAmbito.SelectedValue = "NAC", RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkFCSemanal_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFCSemanal.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCajaSemanal.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Trimestral " & txtFechaFin.Text & ".xls"
        Else
            NombreArchivo = "Flujo Caja Trimestral Exterior " & txtFechaFin.Text & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFCS(rdbAmbito.SelectedValue = "NAC", RutaPlantilla, RutaArchivo, Convert.ToDateTime(txtFechaIni.Text), Convert.ToDateTime(txtFechaFin.Text))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkFCSemanal2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFCSemanal2.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCajaSemanal2.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Semanal " & txtFechaFin.Text & ".xls"
        Else
            NombreArchivo = "Flujo Caja Semanal Exterior " & txtFechaFin.Text & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFCS2(rdbAmbito.SelectedValue = "NAC", False, "MontoBruto", RutaPlantilla, RutaArchivo, Convert.ToDateTime(txtFechaIni.Text), Convert.ToDateTime(txtFechaFin.Text))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkFCSemanal2d_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFCSemanal2d.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCajaSemanal2.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Diario " & txtFechaFin.Text & ".xls"
        Else
            NombreArchivo = "Flujo Caja Diario Exterior " & txtFechaFin.Text & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo

        GeneraReportesFCS2(rdbAmbito.SelectedValue = "NAC", True, "MontoBruto", RutaPlantilla, RutaArchivo, Convert.ToDateTime(txtFechaIni.Text), Convert.ToDateTime(txtFechaFin.Text))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkFCSemanal2b_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFCSemanal2b.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCajaSemanal2.xls"
        If rdbAmbito.SelectedValue = "NAC" Then
            NombreArchivo = "Flujo Caja Semanal RAG " & txtFechaFin.Text & ".xls"
        Else
            NombreArchivo = "Flujo Caja Semanal RAG Exterior " & txtFechaFin.Text & ".xls"
        End If
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFCS2(rdbAmbito.SelectedValue = "NAC", False, "MontoNeto", RutaPlantilla, RutaArchivo, Convert.ToDateTime(txtFechaIni.Text), Convert.ToDateTime(txtFechaFin.Text))
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Protected Sub lnkGenerarReporteN_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGenerarReporteN.Click
        Dim NombrePlantilla, NombreArchivo, RutaPlantilla, RutaArchivo As String
        Dim IdPeriodo As String

        IdPeriodo = cboPeriodo.SelectedValue
        NombrePlantilla = "FlujoCajaN.xls"
        NombreArchivo = "Flujo Caja (Formato B) " & IdPeriodo.Substring(0, 4) & "-" & IdPeriodo.Substring(4, 2) & ".xls"
        RutaPlantilla = MapPath("Plantillas") & "\" + NombrePlantilla
        RutaArchivo = MapPath("Reportes") & "\" + NombreArchivo
        GeneraReportesFCN(RutaPlantilla, RutaArchivo, IdPeriodo)
        GeneraArchivoDescarga(Response, RutaArchivo, NombreArchivo)
    End Sub

    Sub CargaCombos()
        cboPeriodo.DataSource = LlenaPeriodos()
        cboPeriodo.DataTextField = "Periodo"
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataBind()
    End Sub

    'Protected Sub lnkAjuste2010_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuste2010.Click
    '    Dim Fecha As Date

    '    Fecha = Convert.ToDateTime("2010-01-01")
    '    While Fecha <= Convert.ToDateTime("2010-12-31")
    '        CalculaAjustesTipoCambio(Fecha)
    '        Fecha = Fecha.AddDays(1)
    '    End While
    'End Sub

End Class
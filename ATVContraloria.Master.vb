Public Class ATVContraloria
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario As String
        Dim CodTipo As String = ""

        If Session("IdUsuario") Is Nothing Then Response.Redirect("Login.aspx")

        IdUsuario = Session("IdUsuario").ToString
        lblUsuario.Text = FuncionesSeguridad.BuscaUsuario(IdUsuario, CodTipo)
        lblFecha.Text = Date.Now.ToString("dd/MM/yyyy")

        If CodTipo = "ADM" Then
            litMenu.Text = FuncionesSeguridad.GenerarMenu("", "Importación")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu("", "Actualización")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu("", "Mantenimientos")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu("", "Reportes")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu("", "Consultas")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu("", "Seguridad")

            'litImportacion.Text = FuncionesSeguridad.GeneraSubMenu("", "Importación")
            'litActualizacion.Text = FuncionesSeguridad.GeneraSubMenu("", "Actualización")
            'litMantenimientos.Text = FuncionesSeguridad.GeneraSubMenu("", "Mantenimientos")
            'litReportes.Text = FuncionesSeguridad.GeneraSubMenu("", "Reportes")
            'litConsultas.Text = FuncionesSeguridad.GeneraSubMenu("", "Consultas")
            'litSeguridad.Text = FuncionesSeguridad.GeneraSubMenu("", "Seguridad")
        Else

            litMenu.Text = FuncionesSeguridad.GenerarMenu(IdUsuario, "Importación")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu(IdUsuario, "Actualización")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu(IdUsuario, "Mantenimientos")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu(IdUsuario, "Reportes")
            litMenu.Text &= FuncionesSeguridad.GenerarMenu(IdUsuario, "Consultas")
            'litMenu.Text &= ObjWcA.GenerarMenu(IdUsuario, "Seguridad")


            'litImportacion.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Importación")
            'litActualizacion.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Actualización")
            'litMantenimientos.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Mantenimientos")
            'litReportes.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Reportes")
            'litConsultas.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Consultas")
            'litSeguridad.Text = FuncionesSeguridad.GeneraSubMenu(IdUsuario, "Seguridad")
        End If

    End Sub

    Protected Sub lnkSalir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkSalir.Click
        Session("IdUsuario") = Nothing
        Response.Redirect("Login.aspx")
    End Sub

End Class
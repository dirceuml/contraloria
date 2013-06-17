Public Class Aviso
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            lblUsuario.Text = FuncionesSeguridad.BuscaUsuario(IdUsuario, CodTipo)
        End If

        If Not IsPostBack Then
            lblMensaje.Text = "No tiene acceso a la opción seleccionada"
        End If
    End Sub

End Class
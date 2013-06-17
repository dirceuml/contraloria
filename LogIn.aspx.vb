Public Class LogIn
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub lnkIngresar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkIngresar.Click
        Dim IdUsuario As String = ""
        Dim Usuario As String
        Dim CodTipo As String = ""

        If txtCodUsuario.Text = "" Or txtPassword.Text = "" Then
            lblMensaje.Text = "Cód. Usuario y/o Password son obligatorios"
        ElseIf FuncionesSeguridad.ValidaUsuario(txtCodUsuario.Text, txtPassword.Text, IdUsuario) Then
            Session("IdUsuario") = IdUsuario
            Usuario = FuncionesSeguridad.BuscaUsuario(IdUsuario, CodTipo)
            Session("CodTipo") = CodTipo
            Response.Redirect("Principal.aspx")
        Else
            lblMensaje.Text = "Cód. Usuario y/o Password incorrectos"
        End If
    End Sub

End Class
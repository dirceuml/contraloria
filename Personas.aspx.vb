Imports System.Data.OleDb
Imports System.Configuration

Public Class Personas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Personas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        gdvPersona.PageSize = Funciones.iREGxPAG
        LlenaGridPersona(0)
    End Sub

    Sub LlenaGridPersona(ByVal pagina As Integer)
        Dim sSql1, sSql2 As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        'Unimos los filtros
        sSql2 = "From Persona WHERE CodPersona > 0 "
        If txtPersona.Text <> "" Then
            sSql2 += " AND Persona LIKE '%" + txtPersona.Text + "%' "
        End If

        'Extraemos el numero de registros
        hdfNroRegistros.Value = Funciones.ExtraerValor("Select Count(*) as NumReg " + sSql2, "NumReg")
        If hdfNroRegistros.Value <> "0" Then
            lblNumRegistros.Text = "Se han encontrado " + hdfNroRegistros.Value + " registros"
            'LLenamos la grila
            sSql1 = " Select CodPersona, Persona "
            sSql2 += "ORDER BY Persona"

            cn = New OleDbConnection()
            cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
            cn.Open()
            cmd = New OleDbCommand(sSql1 + sSql2, cn)
            dtadap = New OleDbDataAdapter(cmd)
            dtaset = New DataSet()
            dtadap.Fill(dtaset)
            gdvPersona.DataSource = dtaset.Tables(0)
            gdvPersona.PageIndex = pagina
            gdvPersona.DataBind()
            cn.Close()
        Else
            lblNumRegistros.Text = "No se encontro información"
            gdvPersona.DataSource = Nothing
            gdvPersona.DataBind()
        End If
    End Sub

    Protected Sub gdvPersona_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        LlenaGridPersona(e.NewPageIndex)
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click

    End Sub

End Class
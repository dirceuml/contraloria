Imports System.Data.OleDb
Imports System.Configuration

Public Class GrupoATV
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "GrupoATV.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvGrupoATV.PageSize = Funciones.iREGxPAG
            LlenaGrIdEmpresa(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGrIdEmpresa(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select g.IdEmpresa, g.Empresa, g.CodPersona, p.Persona, g.Nivel, g.NivelN, g.Orden "
        sql += "From GrupoATV g, Persona p "
        sql += "Where p.CodPersona = g.CodPersona "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvGrupoATV.DataSource = dtaset.Tables(0)
        gdvGrupoATV.PageIndex = pagina
        gdvGrupoATV.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvGrupoATV_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sIdEmpresa As String = gdvGrupoATV.DataKeys(e.Row.RowIndex).Value.ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCodPersona As DropDownList = CType(e.Row.FindControl("cboCodPersona"), DropDownList)
                LlenarCodPersona(cboCodPersona)
                cboCodPersona.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodPersona").ToString()
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodPersona As DropDownList = CType(e.Row.FindControl("cboCodPersonaNew"), DropDownList)
            LlenarCodPersona(cboCodPersona)
        End If
    End Sub

    Protected Sub gdvGrupoATV_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvGrupoATV.EditIndex = e.NewEditIndex
        LlenaGrIdEmpresa(0)
    End Sub

    Protected Sub gdvGrupoATV_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvGrupoATV.Rows(e.RowIndex)
        Dim sIdEmpresa As String = gdvGrupoATV.DataKeys(e.RowIndex).Value.ToString()
        Dim txtEmpresa As TextBox = DirectCast(row.FindControl("txtEmpresa"), TextBox)
        Dim cboCodPersona As DropDownList = DirectCast(row.FindControl("cboCodPersona"), DropDownList)
        Dim cboNivel As DropDownList = DirectCast(row.FindControl("cboNivel"), DropDownList)
        Dim cboNivelN As DropDownList = DirectCast(row.FindControl("cboNivelN"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        lblMensaje.Text = Funciones.ActualizarGrupoATV(sIdEmpresa, txtEmpresa.Text, cboCodPersona.SelectedValue, cboNivel.SelectedValue, cboNivelN.SelectedValue, txtOrden.Text)
        gdvGrupoATV.EditIndex = -1
        LlenaGrIdEmpresa(0)
    End Sub

    Protected Sub gdvGrupoATV_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdEmpresa As String = gdvGrupoATV.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarGrupoATV(sIdEmpresa)
        LlenaGrIdEmpresa(0)
    End Sub

    Protected Sub gdvGrupoATV_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvGrupoATV.EditIndex = -1
        LlenaGrIdEmpresa(0)
    End Sub

    Private Sub LlenarCodPersona(ByVal cboCodPersona As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodPersona, Persona FROM Persona ORDER BY Persona "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodPersona.DataSource = dtaset.Tables(0)
        cboCodPersona.DataValueField = "CodPersona"
        cboCodPersona.DataTextField = "Persona"
        cboCodPersona.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtEmpresa As TextBox = DirectCast(gdvGrupoATV.FooterRow.FindControl("txtEmpresaNew"), TextBox)
        Dim cboCodPersona As DropDownList = DirectCast(gdvGrupoATV.FooterRow.FindControl("cboCodPersonaNew"), DropDownList)
        Dim cboNivel As DropDownList = DirectCast(gdvGrupoATV.FooterRow.FindControl("cboNivelNew"), DropDownList)
        Dim cboNivelN As DropDownList = DirectCast(gdvGrupoATV.FooterRow.FindControl("cboNivelNNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvGrupoATV.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarGrupoATV(txtEmpresa.Text, cboCodPersona.SelectedValue, cboNivel.SelectedValue, cboNivelN.SelectedValue, txtOrden.Text)
        gdvGrupoATV.EditIndex = -1
        LlenaGrIdEmpresa(0)
    End Sub

End Class
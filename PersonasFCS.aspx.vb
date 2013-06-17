Imports System.Data.OleDb
Imports System.Configuration

Public Class PersonasFCS
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "PersonasFCS.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvPersonaS.PageSize = Funciones.iREGxPAG
            LlenaGridPersonaS(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridPersonaS(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdPersonaS, CodGrupo, PersonaS, CodPersona "
        sql += "From PersonaS "
        sql += "ORDER BY CodGrupo, PersonaS"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvPersonaS.DataSource = dtaset.Tables(0)
        gdvPersonaS.PageIndex = pagina
        gdvPersonaS.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvPersonaS_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCodGrupo As DropDownList = CType(e.Row.FindControl("cboCodGrupo"), DropDownList)
                LlenarCodGrupo(cboCodGrupo)
                cboCodGrupo.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodGrupo").ToString()

                Dim cboCodPersona As DropDownList = CType(e.Row.FindControl("cboCodPersona"), DropDownList)
                LlenarCodPersona(cboCodPersona)
                cboCodPersona.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodPersona").ToString()
            Else
                Dim sCodGrupo As String = DataBinder.Eval(e.Row.DataItem, "CodGrupo").ToString()
                Dim lblGrupo As Label = CType(e.Row.FindControl("lblGrupo"), Label)
                lblGrupo.Text = Funciones.ExtraerValor("Select Grupo From GrupoPersonaS Where CodGrupo = '" + sCodGrupo + "'", "Grupo")

                Dim sCodPersona As String = DataBinder.Eval(e.Row.DataItem, "CodPersona").ToString()
                Dim lblPersona As Label = CType(e.Row.FindControl("lblPersona"), Label)
                If sCodPersona <> "" Then
                    lblPersona.Text = Funciones.ExtraerValor("Select Persona From Persona Where CodPersona = " + sCodPersona, "Persona")
                Else
                    lblPersona.Text = ""
                End If
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
                Dim cboCodGrupoNew As DropDownList = CType(e.Row.FindControl("cboCodGrupoNew"), DropDownList)
                LlenarCodGrupo(cboCodGrupoNew)

                Dim cboCodPersonaNew As DropDownList = CType(e.Row.FindControl("cboCodPersonaNew"), DropDownList)
                LlenarCodPersona(cboCodPersonaNew)
        End If
    End Sub

    Protected Sub gdvPersonaS_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvPersonaS.EditIndex = e.NewEditIndex
        LlenaGridPersonaS(0)
    End Sub

    Protected Sub gdvPersonaS_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvPersonaS.Rows(e.RowIndex)
        Dim sIdPersonaS As String = gdvPersonaS.DataKeys(e.RowIndex).Value.ToString()
        Dim cboCodGrupo As DropDownList = DirectCast(row.FindControl("cboCodGrupo"), DropDownList)
        Dim sPersonaS As TextBox = DirectCast(row.FindControl("txtPersonaS"), TextBox)
        Dim cboCodPersona As DropDownList = DirectCast(row.FindControl("cboCodPersona"), DropDownList)
        lblMensaje.Text = Funciones.ActualizarPersonaS(sIdPersonaS, cboCodGrupo.SelectedValue, sPersonaS.Text, cboCodPersona.SelectedValue)
        gdvPersonaS.EditIndex = -1
        LlenaGridPersonaS(0)
    End Sub

    Protected Sub gdvPersonaS_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvPersonaS.EditIndex = -1
        LlenaGridPersonaS(0)
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboCodGrupo As DropDownList = DirectCast(gdvPersonaS.FooterRow.FindControl("cboCodGrupoNew"), DropDownList)
        Dim sPersonaS As TextBox = DirectCast(gdvPersonaS.FooterRow.FindControl("txtPersonaSNew"), TextBox)
        Dim cboCodPersona As DropDownList = DirectCast(gdvPersonaS.FooterRow.FindControl("cboCodPersonaNew"), DropDownList)
        lblMensaje.Text = Funciones.InsertarPersonaS(cboCodGrupo.SelectedValue, sPersonaS.Text, cboCodPersona.SelectedValue)
        gdvPersonaS.EditIndex = -1
        LlenaGridPersonaS(0)
    End Sub

    Protected Sub gdvPersonaS_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdPersonaS As String = gdvPersonaS.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarPersonaS(sIdPersonaS)
        LlenaGridPersonaS(0)
    End Sub

    Private Sub LlenarCodGrupo(ByVal cboCodGrupo As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodGrupo, Grupo FROM GrupoPersonaS ORDER BY Grupo "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodGrupo.DataSource = dtaset.Tables(0)
        cboCodGrupo.DataValueField = "CodGrupo"
        cboCodGrupo.DataTextField = "Grupo"
        cboCodGrupo.DataBind()
        cn.Close()
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

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        Dim grid As GridView
        grid = gdvPersonaS
        grid.Columns(3).Visible = False
        grid.FooterRow.Visible = False
        FuncionesRep.ExportarExcel(Response, grid, "Personas(Semanal).xls")
    End Sub

End Class
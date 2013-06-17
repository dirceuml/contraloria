Imports System.Data.OleDb
Imports System.Configuration

Public Class GruposPersonas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "GruposPersonas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargarCombos()
            hdfPagina.Value = 0
            gdvPersona.PageSize = Funciones.iREGxPAG
            LlenaGridPersona(0)
        End If
        If hdfAccion.Value = "Recargar" Then
            LlenaGridPersona(0)
            hdfAccion.Value = ""
        End If
    End Sub

    Sub CargarCombos()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = ""
        sql += "select 0 as CodPersonaN, ' -- TODOS LOS GRUPOS -- ' AS PersonaN "
        sql += "UNION "
        sql += "select -1 as CodPersonaN, ' --SIN ASIGNAR-- ' AS PersonaN "
        sql += "UNION SELECT DISTINCT CodPersonaN, PersonaN FROM PersonaN "
        sql += "ORDER BY 2 "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboGrupo.DataSource = dtaset.Tables(0)
        cboGrupo.DataValueField = "CodPersonaN"
        cboGrupo.DataTextField = "PersonaN"
        cboGrupo.DataBind()
        cn.Close()
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
        If cboGrupo.Text = "-1" Then
            sSql2 += " AND CodPersona NOT IN (SELECT PN.CodPersona FROM PersonaN PN, Persona P WHERE PN.CodPersona = P.CodPersona) "
        ElseIf cboGrupo.Text <> "0" Then
            sSql2 += " AND CodPersona IN (SELECT PN.CodPersona FROM PersonaN PN, Persona P WHERE PN.CodPersona = P.CodPersona AND PN.CodPersonaN = " + cboGrupo.Text + ") "
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

    Protected Sub gdvPersona_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sSql As String
        Dim sCodPersona, sPersona As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            'e.Row.Cells(0).Text = (e.Row.DataItemIndex + 1).ToString()

            sCodPersona = DataBinder.Eval(e.Row.DataItem, "CodPersona").ToString()
            sPersona = DataBinder.Eval(e.Row.DataItem, "Persona").ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboPersonaN As DropDownList = CType(e.Row.FindControl("cboPersonaN"), DropDownList)
                LlenarPersonaN(cboPersonaN)
                sSql = "Select DISTINCT CodPersonaN From PersonaN "
                sSql += "Where CodPersona = " + sCodPersona + " "
                Dim CodPersonaN As String = Funciones.ExtraerValor(sSql, "CodPersonaN")
                cboPersonaN.SelectedValue = CodPersonaN
            Else
                sSql = "Select DISTINCT PersonaN From PersonaN "
                sSql += "Where CodPersona = " + sCodPersona + " "
                Funciones.ExtraerValor(sSql, "PersonaN")
                Dim lnkPersonaN As Label = CType(e.Row.FindControl("lnkPersonaN"), Label)
                lnkPersonaN.Text = Funciones.ExtraerValor(sSql, "PersonaN")
            End If
        End If
    End Sub

    Protected Sub gdvPersona_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        LlenaGridPersona(e.NewPageIndex)
        hdfPagina.Value = e.NewPageIndex
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click

    End Sub

    Protected Sub gdvPersona_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvPersona.EditIndex = e.NewEditIndex
        LlenaGridPersona(Convert.ToInt32(hdfPagina.Value))
    End Sub

    Protected Sub gdvPersona_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvPersona.Rows(e.RowIndex)
        Dim cboPersonaN As DropDownList = DirectCast(row.FindControl("cboPersonaN"), DropDownList)
        Dim txtPersonaN As TextBox = DirectCast(row.FindControl("txtPersonaN"), TextBox)
        Dim sCodPersona As String = gdvPersona.DataKeys(e.RowIndex).Value.ToString()
        If cboPersonaN.SelectedValue <> "-1" Then
            Funciones.IngresarCboPersonaN(sCodPersona, cboPersonaN.SelectedValue)
        Else
            Funciones.IngresarTxtPersonaN(sCodPersona, txtPersonaN.Text)
        End If
        gdvPersona.EditIndex = -1
        LlenaGridPersona(Convert.ToInt32(hdfPagina.Value))
    End Sub

    Protected Sub gdvPersona_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvPersona.EditIndex = -1
        LlenaGridPersona(Convert.ToInt32(hdfPagina.Value))
    End Sub

    Private Sub LlenarPersonaN(ByVal cboPersonaN As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "Select -1 as CodPersonaN, '--SIN ASIGNAR--' AS PersonaN "
        sql += "UNION "
        sql += "SELECT DISTINCT CodPersonaN, PersonaN FROM PersonaN ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboPersonaN.DataSource = dtaset.Tables(0)
        cboPersonaN.DataValueField = "CodPersonaN"
        cboPersonaN.DataTextField = "PersonaN"
        cboPersonaN.DataBind()
        cn.Close()
    End Sub

End Class
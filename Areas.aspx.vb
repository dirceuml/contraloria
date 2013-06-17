Imports System.Data.OleDb
Imports System.Configuration

Public Class Areas
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "Areas.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvArea.PageSize = Funciones.iREGxPAG
            Session("OrdenaPor") = ""
            LlenaGridArea(0)
        End If
    End Sub

    Sub LlenaGridArea(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select a.CodCentroCosto, cc.CentroCosto, a.CodArea, a.Area, a.CodEstado, a.FlagPrograma "
        sql += "From CentroCosto cc, Area a "
        sql += "Where a.CodCentroCosto = cc.CodCentroCosto "
        If Session("OrdenaPor").ToString() = "CentroCosto" Then
            sql += "ORDER BY cc.CentroCosto, a.Area"
        ElseIf Session("OrdenaPor").ToString() = "Area" Then
            sql += "ORDER BY a.Area, cc.CentroCosto"
        ElseIf Session("OrdenaPor").ToString() = "CodCentroCosto" Then
            sql += "ORDER BY a.CodCentroCosto"
        ElseIf Session("OrdenaPor").ToString() = "CodArea" Then
            sql += "ORDER BY a.CodArea"
        ElseIf Session("OrdenaPor").ToString() = "CodEstado" Then
            sql += "ORDER BY a.CodEstado"
        Else
            sql += "ORDER BY cc.CentroCosto, a.Area"
        End If
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvArea.DataSource = dtaset.Tables(0)
        gdvArea.PageIndex = pagina
        gdvArea.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvArea_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sFlagPrograma As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                'nada
            Else
                sFlagPrograma = DataBinder.Eval(e.Row.DataItem, "FlagPrograma").ToString()
                If sFlagPrograma = "S" Then
                    sFlagPrograma = "SI"
                Else
                    sFlagPrograma = "NO"
                End If
                Dim lnkPrograma As Label = CType(e.Row.FindControl("lnkPrograma"), Label)
                lnkPrograma.Text = sFlagPrograma
            End If
        End If
    End Sub

    Protected Sub gdvArea_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvArea.EditIndex = e.NewEditIndex
        LlenaGridArea(0)
    End Sub

    Protected Sub gdvArea_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvArea.Rows(e.RowIndex)
        Dim cboPrograma As DropDownList = DirectCast(row.FindControl("cboPrograma"), DropDownList)
        Dim sCodArea As String = gdvArea.DataKeys(e.RowIndex).Value.ToString()
        Funciones.ActualizarArea(sCodArea, cboPrograma.SelectedValue)
        gdvArea.EditIndex = -1
        LlenaGridArea(0)
    End Sub

    Protected Sub gdvArea_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvArea.EditIndex = -1
        LlenaGridArea(0)
    End Sub

    Protected Sub gdvArea_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor") = e.SortExpression
        LlenaGridArea(0)
    End Sub

End Class
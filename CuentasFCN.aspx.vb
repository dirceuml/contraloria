Imports System.Data.OleDb
Imports System.Configuration

Public Class CuentasFCN
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "CuentasFCN.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvCuentaFC.PageSize = Funciones.iREGxPAG
            Session("OrdenaPor1") = ""
            LlenaGridCuentaFC(0)
            Session("OrdenaPor2") = ""
            LlenaGridCuentaFCN(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCuentaFC(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select CodCuenta, Cuenta, CodEstado "
        sql += "From CuentaFC "
        If Session("OrdenaPor1").ToString() = "CodCuenta" Then
            sql += "ORDER BY CodCuenta"
        ElseIf Session("OrdenaPor1").ToString() = "Cuenta" Then
            sql += "ORDER BY Cuenta"
        ElseIf Session("OrdenaPor1").ToString() = "CodEstado" Then
            sql += "ORDER BY CodEstado"
        Else
            sql += "ORDER BY Cuenta"
        End If

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCuentaFC.DataSource = dtaset.Tables(0)
        gdvCuentaFC.PageIndex = pagina
        gdvCuentaFC.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCuentaFC_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor1") = e.SortExpression
        LlenaGridCuentaFC(0)
    End Sub

    Sub LlenaGridCuentaFCN(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select CodCuentaN, CuentaN, CodEstado "
        sql += "From CuentaFCN "
        If Session("OrdenaPor2").ToString() = "CodCuentaN" Then
            sql += "ORDER BY CodCuentaN"
        ElseIf Session("OrdenaPor2").ToString() = "CuentaN" Then
            sql += "ORDER BY CuentaN"
        ElseIf Session("OrdenaPor2").ToString() = "CodEstado" Then
            sql += "ORDER BY CodEstado"
        Else
            sql += "ORDER BY CuentaN"
        End If

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCuentaFCN.DataSource = dtaset.Tables(0)
        gdvCuentaFCN.PageIndex = pagina
        gdvCuentaFCN.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCuentaFCN_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                'nada
            Else
                'sFlagPrograma = DataBinder.Eval(e.Row.DataItem, "FlagPrograma").ToString()
                'If sFlagPrograma = "S" Then
                '    sFlagPrograma = "SI"
                'Else
                '    sFlagPrograma = "NO"
                'End If
                'Dim lnkPrograma As Label = CType(e.Row.FindControl("lnkPrograma"), Label)
                'lnkPrograma.Text = sFlagPrograma
            End If
        End If
    End Sub

    Protected Sub gdvCuentaFCN_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCuentaFCN.EditIndex = e.NewEditIndex
        LlenaGridCuentaFCN(0)
    End Sub

    Protected Sub gdvCuentaFCN_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCuentaFCN.Rows(e.RowIndex)
        Dim sCuentaN As TextBox = DirectCast(row.FindControl("txtCuentaN"), TextBox)
        Dim sCodEstado As TextBox = DirectCast(row.FindControl("txtCodEstado"), TextBox)
        Dim sCodCuentaN As String = gdvCuentaFCN.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.ActualizarCuentaFCN(sCodCuentaN, sCuentaN.Text, sCodEstado.Text)
        gdvCuentaFCN.EditIndex = -1
        LlenaGridCuentaFCN(0)
    End Sub

    Protected Sub gdvCuentaFCN_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sCodCuentaN As String = gdvCuentaFCN.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCuentaFCN(sCodCuentaN)
        LlenaGridCuentaFCN(0)
    End Sub

    Protected Sub gdvCuentaFCN_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCuentaFCN.EditIndex = -1
        LlenaGridCuentaFCN(0)
    End Sub

    Protected Sub gdvCuentaFCN_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor2") = e.SortExpression
        LlenaGridCuentaFCN(0)
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sCuentaNNew As TextBox = DirectCast(gdvCuentaFCN.FooterRow.FindControl("txtCuentaNNew"), TextBox)
        Dim sCodEstadoNew As TextBox = DirectCast(gdvCuentaFCN.FooterRow.FindControl("txtCodEstadoNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCuentaFCN(sCuentaNNew.Text, sCodEstadoNew.Text)
        gdvCuentaFCN.EditIndex = -1
        LlenaGridCuentaFCN(0)
    End Sub

End Class
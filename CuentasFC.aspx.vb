Imports System.Data.OleDb
Imports System.Configuration

Public Class CuentasFC
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "CuentasFC.aspx"
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
            LlenaGridCuentaFC2(0)
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

    Sub LlenaGridCuentaFC2(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select CodCuenta2, Cuenta2, CodEstado "
        sql += "From CuentaFC2 "
        If Session("OrdenaPor2").ToString() = "CodCuenta2" Then
            sql += "ORDER BY CodCuenta2"
        ElseIf Session("OrdenaPor2").ToString() = "Cuenta2" Then
            sql += "ORDER BY Cuenta2"
        ElseIf Session("OrdenaPor2").ToString() = "CodEstado" Then
            sql += "ORDER BY CodEstado"
        Else
            sql += "ORDER BY Cuenta2"
        End If
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCuentaFC2.DataSource = dtaset.Tables(0)
        gdvCuentaFC2.PageIndex = pagina
        gdvCuentaFC2.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCuentaFC_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor1") = e.SortExpression
        LlenaGridCuentaFC(0)
    End Sub

    Protected Sub gdvCuentaFC2_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboCodEstado As DropDownList = CType(e.Row.FindControl("cboCodEstado"), DropDownList)
                LlenarCodEstado(cboCodEstado)
                cboCodEstado.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodEstado").ToString()
            Else
                Dim sCodEstado As String = DataBinder.Eval(e.Row.DataItem, "CodEstado").ToString()
                Dim lblCodEstado As Label = CType(e.Row.FindControl("lblCodEstado"), Label)
                'lblCodEstado.Text = Funciones.ExtraerValor("Select Estado From Estado Where CodEstado = " + sCodEstado, "Estado")
                lblCodEstado.Text = Funciones.ExtraerValor("Select Estado From Estado", "Estado")
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodEstadoNew As DropDownList = CType(e.Row.FindControl("cboCodEstadoNew"), DropDownList)
            LlenarCodEstado(cboCodEstadoNew)
        End If
    End Sub

    Protected Sub gdvCuentaFC2_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCuentaFC2.EditIndex = e.NewEditIndex
        LlenaGridCuentaFC2(0)
    End Sub

    Protected Sub gdvCuentaFC2_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCuentaFC2.Rows(e.RowIndex)
        Dim sCuenta2 As TextBox = DirectCast(row.FindControl("txtCuenta2"), TextBox)
        Dim cboCodEstado As DropDownList = DirectCast(row.FindControl("cboCodEstado"), DropDownList)
        Dim sCodCuenta2 As String = gdvCuentaFC2.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.ActualizarCuentaFC2(sCodCuenta2, sCuenta2.Text, cboCodEstado.SelectedValue)
        gdvCuentaFC2.EditIndex = -1
        LlenaGridCuentaFC2(0)
    End Sub

    Protected Sub gdvCuentaFC2_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sCodCuenta2 As String = gdvCuentaFC2.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCuentaFC2(sCodCuenta2)
        LlenaGridCuentaFC2(0)
    End Sub

    Protected Sub gdvCuentaFC2_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCuentaFC2.EditIndex = -1
        LlenaGridCuentaFC2(0)
    End Sub

    Protected Sub gdvCuentaFC2_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor2") = e.SortExpression
        LlenaGridCuentaFC2(0)
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sCuenta2New As TextBox = DirectCast(gdvCuentaFC2.FooterRow.FindControl("txtCuenta2New"), TextBox)
        Dim cboCodEstado As DropDownList = DirectCast(gdvCuentaFC2.FooterRow.FindControl("cboCodEstadoNew"), DropDownList)
        lblMensaje.Text = Funciones.InsertarCuentaFC2(sCuenta2New.Text, cboCodEstado.SelectedValue)
        gdvCuentaFC2.EditIndex = -1
        LlenaGridCuentaFC2(0)
    End Sub

    Private Sub LlenarCodEstado(ByVal cboCodEstado As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodEstado, Estado FROM Estado ORDER BY Estado "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodEstado.DataSource = dtaset.Tables(0)
        cboCodEstado.DataValueField = "CodEstado"
        cboCodEstado.DataTextField = "Estado"
        cboCodEstado.DataBind()
        cn.Close()
    End Sub

End Class
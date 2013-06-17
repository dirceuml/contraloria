Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCajaS
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "FlujoCajaS.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            gdvCtaFlujoS.PageSize = Funciones.iREGxPAG
            LlenaGridCtaFlujoS(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaFlujoS(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaFlujo, CtaFlujo, Signo, CodSeccion, FlagModif, CodDetalle, Orden "
        sql += "From CtaFlujoS "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaFlujoS.DataSource = dtaset.Tables(0)
        gdvCtaFlujoS.PageIndex = pagina
        gdvCtaFlujoS.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaFlujoS_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sFlagModif, sSigno As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sIdCtaFlujo As String = gdvCtaFlujoS.DataKeys(e.Row.RowIndex).Value.ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                LlenarCodSeccion(cboCodSeccion)
                cboCodSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()
                Dim cboFlagModif As DropDownList = CType(e.Row.FindControl("cboFlagModif"), DropDownList)
                cboFlagModif.SelectedValue = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
            Else
                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                Dim sCodDetalle = DataBinder.Eval(e.Row.DataItem, "CodDetalle").ToString()
                If sCodDetalle = "" Then
                    hlnkDetalle.Visible = False
                Else
                    hlnkDetalle.NavigateUrl = "FlujoCajaSDetalle.aspx?cod=" + sCodDetalle
                End If

                sSigno = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                If sSigno = "1" Then
                    sSigno = " + "
                ElseIf sSigno = "-1" Then
                    sSigno = " - "
                Else
                    sSigno = "   "
                End If
                Dim lnkSigno As Label = CType(e.Row.FindControl("lnkSigno"), Label)
                lnkSigno.Text = sSigno

                sFlagModif = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
                If sFlagModif = "S" Then
                    sFlagModif = "SI"
                Else
                    sFlagModif = "NO"
                End If
                Dim lnkFlagModif As Label = CType(e.Row.FindControl("lnkFlagModif"), Label)
                lnkFlagModif.Text = sFlagModif
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccionNew"), DropDownList)
            LlenarCodSeccion(cboCodSeccion)
        End If
    End Sub

    Protected Sub gdvCtaFlujoS_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaFlujoS.EditIndex = e.NewEditIndex
        LlenaGridCtaFlujoS(0)
    End Sub

    Protected Sub gdvCtaFlujoS_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaFlujoS.Rows(e.RowIndex)
        Dim sIdCtaFlujo As String = gdvCtaFlujoS.DataKeys(e.RowIndex).Value.ToString()
        Dim txtCtaFlujo As TextBox = DirectCast(row.FindControl("txtCtaFlujo"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim txtCodDetalle As TextBox = DirectCast(row.FindControl("txtCodDetalle"), TextBox)
        txtCodDetalle.Text = txtCodDetalle.Text.ToUpper()
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        lblMensaje.Text = Funciones.ActualizarCtaFlujoS(sIdCtaFlujo, txtCtaFlujo.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtCodDetalle.Text, txtOrden.Text)
        gdvCtaFlujoS.EditIndex = -1
        LlenaGridCtaFlujoS(0)
    End Sub

    Protected Sub gdvCtaFlujoS_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaFlujo As String = gdvCtaFlujoS.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaFlujoS(sIdCtaFlujo)
        LlenaGridCtaFlujoS(0)
    End Sub

    Protected Sub gdvCtaFlujoS_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaFlujoS.EditIndex = -1
        LlenaGridCtaFlujoS(0)
    End Sub

    Private Sub LlenarCodSeccion(ByVal cboCodSeccion As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT DISTINCT CodSeccion FROM CtaFlujoS ORDER BY 1 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodSeccion.DataSource = dtaset.Tables(0)
        cboCodSeccion.DataValueField = "CodSeccion"
        cboCodSeccion.DataTextField = "CodSeccion"
        cboCodSeccion.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaFlujo As TextBox = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("txtCtaFlujoNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("cboCodSeccionNew"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("cboFlagModifNew"), DropDownList)
        Dim txtCodDetalle As TextBox = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("txtCodDetalleNew"), TextBox)
        txtCodDetalle.Text = txtCodDetalle.Text.ToUpper()
        Dim txtOrden As TextBox = DirectCast(gdvCtaFlujoS.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaFlujoS(txtCtaFlujo.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtCodDetalle.Text, txtOrden.Text)
        gdvCtaFlujoS.EditIndex = -1
        LlenaGridCtaFlujoS(0)
    End Sub

End Class
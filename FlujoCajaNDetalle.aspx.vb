Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCajaNDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "FlujoCajaN.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("id") = "" Then Response.Redirect("FlujoCajaN.aspx")

        If Not IsPostBack Then
            hdfIdCtaFlujoNRep.Value = Page.Request.QueryString("id")
            gdvCtaDetalleNRep.PageSize = Funciones.iREGxPAG
            LlenaGridCtaDetalleNRep(0)
            lblCtaFlujoNRep.Text = Funciones.ExtraerValor("Select CtaFlujoNRep From CtaFlujoNRep Where IdCtaFlujoNRep =" + hdfIdCtaFlujoNRep.Value, "CtaFlujoNRep")
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaDetalleNRep(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaDetalleNRep, CtaDetalleNRep, Signo, CodCtaOrigen, Orden, TipoDetalle "
        sql += "From CtaDetalleNRep "
        sql += "Where IdCtaFlujoNRep = " + hdfIdCtaFlujoNRep.Value + " "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaDetalleNRep.DataSource = dtaset.Tables(0)
        gdvCtaDetalleNRep.PageIndex = pagina
        gdvCtaDetalleNRep.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaDetalleNRep_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sSigno, sCodCtaOrigen, sTipoDetalle As String

        If e.Row.RowType = DataControlRowType.DataRow Then

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigen"), DropDownList)
                LlenarCodCtaOrigen(cboCodCtaOrigen)
                cboCodCtaOrigen.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
                Dim cboTipoDetalle As DropDownList = CType(e.Row.FindControl("cboTipoDetalle"), DropDownList)
                cboTipoDetalle.SelectedValue = DataBinder.Eval(e.Row.DataItem, "TipoDetalle").ToString()
            Else
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

                sCodCtaOrigen = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
                Dim lnkCodCtaOrigen As Label = CType(e.Row.FindControl("lnkCodCtaOrigen"), Label)
                lnkCodCtaOrigen.Text = Funciones.ExtraerValor("Select CuentaN From V_CuentaFCN Where CodCuentaN = '" + sCodCtaOrigen + "'", "CuentaN")

                sTipoDetalle = DataBinder.Eval(e.Row.DataItem, "TipoDetalle").ToString()
                If sTipoDetalle = "ARE" Then
                    sTipoDetalle = "Area"
                ElseIf sTipoDetalle = "GRP" Then
                    sTipoDetalle = "Grupo"
                ElseIf sTipoDetalle = "PER" Then
                    sTipoDetalle = "Persona"
                ElseIf sTipoDetalle = "PRG" Then
                    sTipoDetalle = "Programa"
                Else
                    sTipoDetalle = " "
                End If
                Dim lnkTipoDetalle As Label = CType(e.Row.FindControl("lnkTipoDetalle"), Label)
                lnkTipoDetalle.Text = sTipoDetalle
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigenNew"), DropDownList)
            LlenarCodCtaOrigen(cboCodCtaOrigen)
        End If
    End Sub

    Protected Sub gdvCtaDetalleNRep_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaDetalleNRep.EditIndex = e.NewEditIndex
        LlenaGridCtaDetalleNRep(0)
    End Sub

    Protected Sub gdvCtaDetalleNRep_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaDetalleNRep.Rows(e.RowIndex)
        Dim sIdCtaDetalleNRep As String = gdvCtaDetalleNRep.DataKeys(e.RowIndex).Value.ToString()
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        Dim txtCtaDetalleNRep As TextBox = DirectCast(row.FindControl("txtCtaDetalleNRep"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(row.FindControl("cboTipoDetalle"), DropDownList)
        lblMensaje.Text = Funciones.ActualizarCtaDetalleNRep(sIdCtaDetalleNRep, txtCtaDetalleNRep.Text, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue, txtOrden.Text)
        gdvCtaDetalleNRep.EditIndex = -1
        LlenaGridCtaDetalleNRep(0)
    End Sub

    Protected Sub gdvCtaDetalleNRep_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaDetalleNRep As String = gdvCtaDetalleNRep.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaDetalleNRep(sIdCtaDetalleNRep)
        LlenaGridCtaDetalleNRep(0)
    End Sub

    Protected Sub gdvCtaDetalleNRep_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaDetalleNRep.EditIndex = -1
        LlenaGridCtaDetalleNRep(0)
    End Sub

    Private Sub LlenarCodCtaOrigen(ByVal cboCodCtaOrigen As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT CodCuentaN, CuentaN FROM V_CuentaFCN ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCtaOrigen.DataSource = dtaset.Tables(0)
        cboCodCtaOrigen.DataValueField = "CodCuentaN"
        cboCodCtaOrigen.DataTextField = "CuentaN"
        cboCodCtaOrigen.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaDetalleNRep As TextBox = DirectCast(gdvCtaDetalleNRep.FooterRow.FindControl("txtCtaDetalleNRepNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaDetalleNRep.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaDetalleNRep.FooterRow.FindControl("cboCodCtaOrigenNew"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(gdvCtaDetalleNRep.FooterRow.FindControl("cboTipoDetalleNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaDetalleNRep.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaDetalleNRep(hdfIdCtaFlujoNRep.Value, txtCtaDetalleNRep.Text, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue, txtOrden.Text)
        gdvCtaDetalleNRep.EditIndex = -1
        LlenaGridCtaDetalleNRep(0)
    End Sub

End Class
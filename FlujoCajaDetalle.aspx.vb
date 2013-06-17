Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCajaDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "FlujoCaja.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("id") = "" Then Response.Redirect("FlujoCaja.aspx")

        If Not IsPostBack Then
            hdfIdCtaFlujoRep.Value = Page.Request.QueryString("id")
            gdvCtaDetalleRep.PageSize = Funciones.iREGxPAG
            LlenaGridCtaDetalleRep(0)
            lblCtaFlujoRep.Text = Funciones.ExtraerValor("Select CtaFlujoRep From CtaFlujoRep Where IdCtaFlujoRep =" + hdfIdCtaFlujoRep.Value, "CtaFlujoRep")
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaDetalleRep(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaDetalleRep, CtaDetalleRep, Signo, CodCtaOrigen, Orden, TipoDetalle "
        sql += "From CtaDetalleRep "
        sql += "Where IdCtaFlujoRep = " + hdfIdCtaFlujoRep.Value + " "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaDetalleRep.DataSource = dtaset.Tables(0)
        gdvCtaDetalleRep.PageIndex = pagina
        gdvCtaDetalleRep.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaDetalleRep_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
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
                lnkCodCtaOrigen.Text = Funciones.ExtraerValor("Select Cuenta2 + ' [' + convert(varchar(3), CodCuenta2) + ']' as Cuenta2 From V_CuentaFC Where CodCuenta2 = '" + sCodCtaOrigen + "'", "Cuenta2")

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

    Protected Sub gdvCtaDetalleRep_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaDetalleRep.EditIndex = e.NewEditIndex
        LlenaGridCtaDetalleRep(0)
    End Sub

    Protected Sub gdvCtaDetalleRep_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaDetalleRep.Rows(e.RowIndex)
        Dim sIdCtaDetalleRep As String = gdvCtaDetalleRep.DataKeys(e.RowIndex).Value.ToString()
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        Dim txtCtaDetalleRep As TextBox = DirectCast(row.FindControl("txtCtaDetalleRep"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(row.FindControl("cboTipoDetalle"), DropDownList)
        lblMensaje.Text = Funciones.ActualizarCtaDetalleRep(sIdCtaDetalleRep, txtCtaDetalleRep.Text, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue, txtOrden.Text)
        gdvCtaDetalleRep.EditIndex = -1
        LlenaGridCtaDetalleRep(0)
    End Sub

    Protected Sub gdvCtaDetalleRep_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaDetalleRep As String = gdvCtaDetalleRep.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaDetalleRep(sIdCtaDetalleRep)
        LlenaGridCtaDetalleRep(0)
    End Sub

    Protected Sub gdvCtaDetalleRep_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaDetalleRep.EditIndex = -1
        LlenaGridCtaDetalleRep(0)
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

        sql = " SELECT CodCuenta2, Cuenta2 + ' [' + convert(varchar(3), CodCuenta2) + ']' as Cuenta2 FROM V_CuentaFC ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCtaOrigen.DataSource = dtaset.Tables(0)
        cboCodCtaOrigen.DataValueField = "CodCuenta2"
        cboCodCtaOrigen.DataTextField = "Cuenta2"
        cboCodCtaOrigen.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaDetalleRep As TextBox = DirectCast(gdvCtaDetalleRep.FooterRow.FindControl("txtCtaDetalleRepNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaDetalleRep.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaDetalleRep.FooterRow.FindControl("cboCodCtaOrigenNew"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(gdvCtaDetalleRep.FooterRow.FindControl("cboTipoDetalleNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaDetalleRep.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaDetalleRep(hdfIdCtaFlujoRep.Value, txtCtaDetalleRep.Text, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue, txtOrden.Text)
        gdvCtaDetalleRep.EditIndex = -1
        LlenaGridCtaDetalleRep(0)
    End Sub

End Class
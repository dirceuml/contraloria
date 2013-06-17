Imports System.Data.OleDb
Imports System.Configuration

Public Class EGPDetalle
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "EGP.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Request.QueryString("IdCtaEGP") = "" Then Response.Redirect("EGP.aspx")

        If Not IsPostBack Then
            hdfIdCtaEGP.Value = Page.Request.QueryString("IdCtaEGP")
            lblCtaEGP.Text = Funciones.ExtraerValor("Select CtaEGP From CtaEGP Where IdCtaEGP = '" + hdfIdCtaEGP.Value + "'", "CtaEGP")
            gdvCtaEGPDet.PageSize = Funciones.iREGxPAG
            Session("OrdenaPor") = ""
            LlenaGrIdCtaEGPDet(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGrIdCtaEGPDet(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select c.IdCtaEGPDet, c.IdCtaEGP, c.Signo, c.CodCtaOrigen, x.Cuenta, c.TipoDetalle "
        sql += "From CtaEGPDet c, CuentaEGP x "
        sql += "Where c.CodCtaOrigen = x.CodCuenta AND c.IdCtaEGP = '" + hdfIdCtaEGP.Value + "' "
        If Session("OrdenaPor").ToString() = "Cuenta" Then
            sql += "ORDER BY x.Cuenta"
        Else
            sql += "ORDER BY c.CodCtaOrigen"
        End If
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        If (dtaset.Tables(0).Rows.Count = 0) Then
            Dim dRow As DataRow
            dRow = dtaset.Tables(0).NewRow()
            dRow(0) = "0"
            dtaset.Tables(0).Rows.Add(dRow)
        End If
        gdvCtaEGPDet.DataSource = dtaset.Tables(0)
        gdvCtaEGPDet.PageIndex = pagina
        gdvCtaEGPDet.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaEGPDet_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
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
                If sSigno = "" Then
                    e.Row.Visible = False
                End If
                If sSigno = "1" Then
                    sSigno = " + "
                ElseIf sSigno = "-1" Then
                    sSigno = " - "
                End If
                Dim lnkSigno As Label = CType(e.Row.FindControl("lnkSigno"), Label)
                lnkSigno.Text = sSigno

                sCodCtaOrigen = DataBinder.Eval(e.Row.DataItem, "CodCtaOrigen").ToString()
                Dim lblCodCtaOrigen As Label = CType(e.Row.FindControl("lblCodCtaOrigen"), Label)
                lblCodCtaOrigen.Text = Funciones.ExtraerValor("Select Cuenta From CuentaEGP Where CodCuenta = '" + sCodCtaOrigen + "'", "Cuenta")

                sTipoDetalle = DataBinder.Eval(e.Row.DataItem, "TipoDetalle").ToString()
                If sTipoDetalle = "GRP" Then
                    sTipoDetalle = "Grupo"
                ElseIf sTipoDetalle = "PER" Then
                    sTipoDetalle = "Persona"
                Else
                    sTipoDetalle = " "
                End If
                Dim lblTipoDetalle As Label = CType(e.Row.FindControl("lblTipoDetalle"), Label)
                lblTipoDetalle.Text = sTipoDetalle
            End If
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            Dim cboCodCtaOrigen As DropDownList = CType(e.Row.FindControl("cboCodCtaOrigenNew"), DropDownList)
            LlenarCodCtaOrigen(cboCodCtaOrigen)
        End If
    End Sub

    Protected Sub gdvCtaEGPDet_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaEGPDet.EditIndex = e.NewEditIndex
        LlenaGrIdCtaEGPDet(0)
    End Sub

    Protected Sub gdvCtaEGPDet_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaEGPDet.Rows(e.RowIndex)
        Dim sIdCtaEGPDet As String = gdvCtaEGPDet.DataKeys(e.RowIndex).Value.ToString()
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(row.FindControl("cboCodCtaOrigen"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(row.FindControl("cboTipoDetalle"), DropDownList)
        lblMensaje.Text = Funciones.ActualizarCtaEGPDet(sIdCtaEGPDet, hdfIdCtaEGP.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue)
        gdvCtaEGPDet.EditIndex = -1
        LlenaGrIdCtaEGPDet(0)
    End Sub

    Protected Sub gdvCtaEGPDet_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaEGPDet As String = gdvCtaEGPDet.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaEGPDet(sIdCtaEGPDet)
        LlenaGrIdCtaEGPDet(0)
    End Sub

    Protected Sub gdvCtaEGPDet_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaEGPDet.EditIndex = -1
        LlenaGrIdCtaEGPDet(0)
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

        sql = " SELECT CodCuenta, '[' + CodCuenta + '] ' + Cuenta AS Cuenta FROM CuentaEGP ORDER BY 2 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodCtaOrigen.DataSource = dtaset.Tables(0)
        cboCodCtaOrigen.DataValueField = "CodCuenta"
        cboCodCtaOrigen.DataTextField = "Cuenta"
        cboCodCtaOrigen.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaEGPDet.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodCtaOrigen As DropDownList = DirectCast(gdvCtaEGPDet.FooterRow.FindControl("cboCodCtaOrigenNew"), DropDownList)
        Dim cboTipoDetalle As DropDownList = DirectCast(gdvCtaEGPDet.FooterRow.FindControl("cboTipoDetalleNew"), DropDownList)
        lblMensaje.Text = Funciones.InsertarCtaEGPDet(hdfIdCtaEGP.Value, cboSigno.SelectedValue, cboCodCtaOrigen.SelectedValue, cboTipoDetalle.SelectedValue)
        gdvCtaEGPDet.EditIndex = -1
        LlenaGrIdCtaEGPDet(0)
    End Sub

    Protected Sub gdvCtaEGPDet_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Session("OrdenaPor") = e.SortExpression
        LlenaGrIdCtaEGPDet(0)
    End Sub

End Class
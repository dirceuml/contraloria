Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCaja
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

        If Not IsPostBack Then
            gdvCtaFlujoRep.PageSize = Funciones.iREGxPAG
            LlenaGridCtaFlujoRep(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaFlujoRep(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaFlujoRep, CtaFlujoRep, Signo, CodSeccion, FlagModif, CodAuxiliar, Orden "
        sql += "From CtaFlujoRep "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaFlujoRep.DataSource = dtaset.Tables(0)
        gdvCtaFlujoRep.PageIndex = pagina
        gdvCtaFlujoRep.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaFlujoRep_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sFlagModif, sSigno As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sIdCtaFlujoRep As String = gdvCtaFlujoRep.DataKeys(e.Row.RowIndex).Value.ToString()

            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim cboSigno As DropDownList = CType(e.Row.FindControl("cboSigno"), DropDownList)
                cboSigno.SelectedValue = DataBinder.Eval(e.Row.DataItem, "Signo").ToString()
                Dim cboCodSeccion As DropDownList = CType(e.Row.FindControl("cboCodSeccion"), DropDownList)
                LlenarCodSeccion(cboCodSeccion)
                cboCodSeccion.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodSeccion").ToString()
                Dim cboFlagModif As DropDownList = CType(e.Row.FindControl("cboFlagModif"), DropDownList)
                cboFlagModif.SelectedValue = DataBinder.Eval(e.Row.DataItem, "FlagModif").ToString()
                Dim cboCodAuxiliar As DropDownList = CType(e.Row.FindControl("cboCodAuxiliar"), DropDownList)
                LlenarCodAuxiliar(cboCodAuxiliar)
                cboCodAuxiliar.SelectedValue = DataBinder.Eval(e.Row.DataItem, "CodAuxiliar").ToString()
            Else
                Dim hlnkDetalle As HyperLink = CType(e.Row.FindControl("hlnkDetalle"), HyperLink)
                hlnkDetalle.NavigateUrl = "FlujoCajaDetalle.aspx?id=" + sIdCtaFlujoRep
                Dim sCodAuxiliar = DataBinder.Eval(e.Row.DataItem, "CodAuxiliar").ToString()
                If sCodAuxiliar = "" Then
                    hlnkDetalle.Visible = False
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
            Dim cboCodAuxiliar As DropDownList = CType(e.Row.FindControl("cboCodAuxiliarNew"), DropDownList)
            LlenarCodAuxiliar(cboCodAuxiliar)
        End If
    End Sub

    Protected Sub gdvCtaFlujoRep_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaFlujoRep.EditIndex = e.NewEditIndex
        LlenaGridCtaFlujoRep(0)
    End Sub

    Protected Sub gdvCtaFlujoRep_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaFlujoRep.Rows(e.RowIndex)
        Dim sIdCtaFlujoRep As String = gdvCtaFlujoRep.DataKeys(e.RowIndex).Value.ToString()
        Dim txtCtaFlujoRep As TextBox = DirectCast(row.FindControl("txtCtaFlujoRep"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim cboCodAuxiliar As DropDownList = DirectCast(row.FindControl("cboCodAuxiliar"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        lblMensaje.Text = Funciones.ActualizarCtaFlujoRep(sIdCtaFlujoRep, txtCtaFlujoRep.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, cboCodAuxiliar.SelectedValue, txtOrden.Text)
        gdvCtaFlujoRep.EditIndex = -1
        LlenaGridCtaFlujoRep(0)
    End Sub

    Protected Sub gdvCtaFlujoRep_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaFlujoRep As String = gdvCtaFlujoRep.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaFlujoRep(sIdCtaFlujoRep)
        LlenaGridCtaFlujoRep(0)
    End Sub

    Protected Sub gdvCtaFlujoRep_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaFlujoRep.EditIndex = -1
        LlenaGridCtaFlujoRep(0)
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

        sql = " SELECT DISTINCT CodSeccion FROM CtaFlujoRep ORDER BY 1 "
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

    Private Sub LlenarCodAuxiliar(ByVal cboCodAuxiliar As DropDownList)
        Dim sql As String = ""
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " SELECT DISTINCT CodAuxiliar FROM CtaFlujoRep ORDER BY 1 "
        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboCodAuxiliar.DataSource = dtaset.Tables(0)
        cboCodAuxiliar.DataValueField = "CodAuxiliar"
        cboCodAuxiliar.DataTextField = "CodAuxiliar"
        cboCodAuxiliar.DataBind()
        cn.Close()
    End Sub

    Protected Sub btnNuevo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtCtaFlujoRep As TextBox = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("txtCtaFlujoRepNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("cboCodSeccionNew"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("cboFlagModifNew"), DropDownList)
        Dim cboCodAuxiliar As DropDownList = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("cboCodAuxiliarNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaFlujoRep.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaFlujoRep(txtCtaFlujoRep.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, cboCodAuxiliar.SelectedValue, txtOrden.Text)
        gdvCtaFlujoRep.EditIndex = -1
        LlenaGridCtaFlujoRep(0)
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        gdvCtaFCExp.DataSource = FuncionesFC.LlenaFCExp
        gdvCtaFCExp.DataBind()
        gdvCtaFCExp.Visible = True
        FuncionesVarias.DescargaExcel(Response, gdvCtaFCExp, "Ctas Flujo Caja Mensual.xls")
    End Sub

End Class
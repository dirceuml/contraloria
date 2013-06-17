Imports System.Data.OleDb
Imports System.Configuration

Public Class EGP
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

        If Not IsPostBack Then
            gdvCtaEGP.PageSize = Funciones.iREGxPAG
            LlenaGridCtaEGP(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaEGP(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaEGP, CtaEGP, Signo, CodSeccion, FlagModif, Orden "
        sql += "From CtaEGP "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaEGP.DataSource = dtaset.Tables(0)
        gdvCtaEGP.PageIndex = pagina
        gdvCtaEGP.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaEGP_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sFlagModif, sSigno As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sIdCtaEGP As String = gdvCtaEGP.DataKeys(e.Row.RowIndex).Value.ToString()

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
                hlnkDetalle.NavigateUrl = "EGPDetalle.aspx?IdCtaEGP=" + sIdCtaEGP
                
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

    Protected Sub gdvCtaEGP_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaEGP.EditIndex = e.NewEditIndex
        LlenaGridCtaEGP(0)
    End Sub

    Protected Sub gdvCtaEGP_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaEGP.Rows(e.RowIndex)
        Dim sIdCtaEGP As String = gdvCtaEGP.DataKeys(e.RowIndex).Value.ToString()
        Dim txtCtaEGP As TextBox = DirectCast(row.FindControl("txtCtaEGP"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        lblMensaje.Text = Funciones.ActualizarCtaEGP(sIdCtaEGP, txtCtaEGP.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaEGP.EditIndex = -1
        LlenaGridCtaEGP(0)
    End Sub

    Protected Sub gdvCtaEGP_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaEGP As String = gdvCtaEGP.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaEGP(sIdCtaEGP)
        LlenaGridCtaEGP(0)
    End Sub

    Protected Sub gdvCtaEGP_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaEGP.EditIndex = -1
        LlenaGridCtaEGP(0)
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

        sql = " SELECT DISTINCT CodSeccion FROM CtaEGP ORDER BY 1 "
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
        Dim txtCtaEGP As TextBox = DirectCast(gdvCtaEGP.FooterRow.FindControl("txtCtaEGPNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaEGP.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaEGP.FooterRow.FindControl("cboCodSeccionNew"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaEGP.FooterRow.FindControl("cboFlagModifNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaEGP.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaEGP(txtCtaEGP.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, txtOrden.Text)
        gdvCtaEGP.EditIndex = -1
        LlenaGridCtaEGP(0)
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDescargaExcel.Click
        gdvCtaEGPExp.DataSource = FuncionesEGP.LlenaEGPExp
        gdvCtaEGPExp.DataBind()
        gdvCtaEGPExp.Visible = True
        FuncionesVarias.DescargaExcel(Response, gdvCtaEGPExp, "Ctas EGP.xls")
    End Sub

End Class
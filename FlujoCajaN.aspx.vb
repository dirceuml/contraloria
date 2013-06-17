Imports System.Data.OleDb
Imports System.Configuration

Public Class FlujoCajaN
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

        If Not IsPostBack Then
            gdvCtaFlujoNRep.PageSize = Funciones.iREGxPAG
            LlenaGridCtaFlujoNRep(0)
            lblMensaje.Text = ""
        End If
    End Sub

    Sub LlenaGridCtaFlujoNRep(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select IdCtaFlujoNRep, CtaFlujoNRep, Signo, CodSeccion, FlagModif, CodAuxiliar, Orden "
        sql += "From CtaFlujoNRep "
        sql += "ORDER BY Orden"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCtaFlujoNRep.DataSource = dtaset.Tables(0)
        gdvCtaFlujoNRep.PageIndex = pagina
        gdvCtaFlujoNRep.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCtaFlujoNRep_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        Dim sFlagModif, sSigno As String

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sIdCtaFlujoNRep As String = gdvCtaFlujoNRep.DataKeys(e.Row.RowIndex).Value.ToString()

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
                hlnkDetalle.NavigateUrl = "FlujoCajaNDetalle.aspx?id=" + sIdCtaFlujoNRep
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
        End If
    End Sub

    Protected Sub gdvCtaFlujoNRep_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvCtaFlujoNRep.EditIndex = e.NewEditIndex
        LlenaGridCtaFlujoNRep(0)
    End Sub

    Protected Sub gdvCtaFlujoNRep_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvCtaFlujoNRep.Rows(e.RowIndex)
        Dim sIdCtaFlujoNRep As String = gdvCtaFlujoNRep.DataKeys(e.RowIndex).Value.ToString()
        Dim txtCtaFlujoNRep As TextBox = DirectCast(row.FindControl("txtCtaFlujoNRep"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(row.FindControl("cboSigno"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(row.FindControl("cboCodSeccion"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(row.FindControl("cboFlagModif"), DropDownList)
        Dim cboCodAuxiliar As DropDownList = DirectCast(row.FindControl("cboCodAuxiliar"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(row.FindControl("txtOrden"), TextBox)
        lblMensaje.Text = Funciones.ActualizarCtaFlujoNRep(sIdCtaFlujoNRep, txtCtaFlujoNRep.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, cboCodAuxiliar.SelectedValue, txtOrden.Text)
        gdvCtaFlujoNRep.EditIndex = -1
        LlenaGridCtaFlujoNRep(0)
    End Sub

    Protected Sub gdvCtaFlujoRep_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim sIdCtaFlujoNRep As String = gdvCtaFlujoNRep.DataKeys(e.RowIndex).Value.ToString()
        lblMensaje.Text = Funciones.BorrarCtaFlujoNRep(sIdCtaFlujoNRep)
        LlenaGridCtaFlujoNRep(0)
    End Sub

    Protected Sub gdvCtaFlujoNRep_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvCtaFlujoNRep.EditIndex = -1
        LlenaGridCtaFlujoNRep(0)
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

        sql = " SELECT DISTINCT CodSeccion FROM CtaFlujoNRep ORDER BY 1 "
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

        sql = " SELECT DISTINCT CodAuxiliar FROM CtaFlujoNRep ORDER BY 1 "
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
        Dim txtCtaFlujoNRep As TextBox = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("txtCtaFlujoNRepNew"), TextBox)
        Dim cboSigno As DropDownList = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("cboSignoNew"), DropDownList)
        Dim cboCodSeccion As DropDownList = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("cboCodSeccionNew"), DropDownList)
        Dim cboFlagModif As DropDownList = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("cboFlagModifNew"), DropDownList)
        Dim cboCodAuxiliar As DropDownList = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("cboCodAuxiliarNew"), DropDownList)
        Dim txtOrden As TextBox = DirectCast(gdvCtaFlujoNRep.FooterRow.FindControl("txtOrdenNew"), TextBox)
        lblMensaje.Text = Funciones.InsertarCtaFlujoNRep(txtCtaFlujoNRep.Text, cboSigno.SelectedValue, cboCodSeccion.SelectedValue, cboFlagModif.SelectedValue, cboCodAuxiliar.SelectedValue, txtOrden.Text)
        gdvCtaFlujoNRep.EditIndex = -1
        LlenaGridCtaFlujoNRep(0)
    End Sub

End Class
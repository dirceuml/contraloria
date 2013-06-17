Imports System.Data.OleDb
Imports System.Configuration

Public Class Presupuesto
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CargarCombos()
            gdvFlujo.PageSize = Funciones.iREGxPAG
            LlenaGridFlujo(0)
            lblMensaje.Text = ""
        End If
        If hdfAccion.Value = "Recargar" Then
            LlenaGridFlujo(0)
            hdfAccion.Value = ""
        End If

    End Sub

    Sub CargarCombos()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "SELECT DISTINCT SUBSTRING(convert(varchar,IdPeriodo),1,4) as IdPeriodo FROM Periodo where IdPeriodo >= 201101"
        sql += "ORDER BY 1 "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)

        cboPeriodo.DataSource = dtaset.Tables(0)
        cboPeriodo.DataValueField = "IdPeriodo"
        cboPeriodo.DataTextField = "IdPeriodo"
        cboPeriodo.DataBind()

        cn.Close()
    End Sub

    Sub LlenaGridFlujo(ByVal pagina As Integer)
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = " Select SUBSTRING(convert(varchar,IdPeriodo),1,4) as IdPeriodo, Orden, CodSeccion, CodDetalle, Rubro "
        sql += "From V_FlujoS Where CodVersion = 'P' and IdPeriodo = " + cboPeriodo.Text + "01 "
        sql += "ORDER BY Orden "

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvFlujo.DataSource = dtaset.Tables(0)
        gdvFlujo.PageIndex = pagina
        gdvFlujo.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvFlujo_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sCodDetalle As String = DataBinder.Eval(e.Row.DataItem, "CodDetalle").ToString()
            Dim sRubro As String = DataBinder.Eval(e.Row.DataItem, "Rubro").ToString()
            Dim sSql = "Select ISNULL(sum(MontoNeto),0) as MontoNeto From DetalleS Where CodVersion = 'P' AND CodDetalle = '" + sCodDetalle + "' AND Rubro = '" + sRubro + "' AND IdPeriodo = " + cboPeriodo.Text
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim hdfRubro As HiddenField = CType(e.Row.FindControl("hdfRubro"), HiddenField)
                hdfRubro.value = sRubro
                Dim txtMontoNetoEnero As TextBox = CType(e.Row.FindControl("txtMontoNetoEnero"), TextBox)
                txtMontoNetoEnero.Text = Funciones.ExtraerValor(sSql + "01", "MontoNeto")
                Dim txtMontoNetoFebrero As TextBox = CType(e.Row.FindControl("txtMontoNetoFebrero"), TextBox)
                txtMontoNetoFebrero.Text = Funciones.ExtraerValor(sSql + "02", "MontoNeto")
                Dim txtMontoNetoMarzo As TextBox = CType(e.Row.FindControl("txtMontoNetoMarzo"), TextBox)
                txtMontoNetoMarzo.Text = Funciones.ExtraerValor(sSql + "03", "MontoNeto")
                Dim txtMontoNetoAbril As TextBox = CType(e.Row.FindControl("txtMontoNetoAbril"), TextBox)
                txtMontoNetoAbril.Text = Funciones.ExtraerValor(sSql + "04", "MontoNeto")
                Dim txtMontoNetoMayo As TextBox = CType(e.Row.FindControl("txtMontoNetoMayo"), TextBox)
                txtMontoNetoMayo.Text = Funciones.ExtraerValor(sSql + "05", "MontoNeto")
                Dim txtMontoNetoJunio As TextBox = CType(e.Row.FindControl("txtMontoNetoJunio"), TextBox)
                txtMontoNetoJunio.Text = Funciones.ExtraerValor(sSql + "06", "MontoNeto")
                Dim txtMontoNetoJulio As TextBox = CType(e.Row.FindControl("txtMontoNetoJulio"), TextBox)
                txtMontoNetoJulio.Text = Funciones.ExtraerValor(sSql + "07", "MontoNeto")
                Dim txtMontoNetoAgosto As TextBox = CType(e.Row.FindControl("txtMontoNetoAgosto"), TextBox)
                txtMontoNetoAgosto.Text = Funciones.ExtraerValor(sSql + "08", "MontoNeto")
                Dim txtMontoNetoSetiembre As TextBox = CType(e.Row.FindControl("txtMontoNetoSetiembre"), TextBox)
                txtMontoNetoSetiembre.Text = Funciones.ExtraerValor(sSql + "09", "MontoNeto")
                Dim txtMontoNetoOctubre As TextBox = CType(e.Row.FindControl("txtMontoNetoOctubre"), TextBox)
                txtMontoNetoOctubre.Text = Funciones.ExtraerValor(sSql + "10", "MontoNeto")
                Dim txtMontoNetoNoviembre As TextBox = CType(e.Row.FindControl("txtMontoNetoNoviembre"), TextBox)
                txtMontoNetoNoviembre.Text = Funciones.ExtraerValor(sSql + "11", "MontoNeto")
                Dim txtMontoNetoDiciembre As TextBox = CType(e.Row.FindControl("txtMontoNetoDiciembre"), TextBox)
                txtMontoNetoDiciembre.Text = Funciones.ExtraerValor(sSql + "12", "MontoNeto")
            Else
                Dim lblMontoNetoEnero As Label = CType(e.Row.FindControl("lblMontoNetoEnero"), Label)
                lblMontoNetoEnero.Text = Funciones.ExtraerValor(sSql + "01", "MontoNeto")
                Dim lblMontoNetoFebrero As Label = CType(e.Row.FindControl("lblMontoNetoFebrero"), Label)
                lblMontoNetoFebrero.Text = Funciones.ExtraerValor(sSql + "02", "MontoNeto")
                Dim lblMontoNetoMarzo As Label = CType(e.Row.FindControl("lblMontoNetoMarzo"), Label)
                lblMontoNetoMarzo.Text = Funciones.ExtraerValor(sSql + "03", "MontoNeto")
                Dim lblMontoNetoAbril As Label = CType(e.Row.FindControl("lblMontoNetoAbril"), Label)
                lblMontoNetoAbril.Text = Funciones.ExtraerValor(sSql + "04", "MontoNeto")
                Dim lblMontoNetoMayo As Label = CType(e.Row.FindControl("lblMontoNetoMayo"), Label)
                lblMontoNetoMayo.Text = Funciones.ExtraerValor(sSql + "05", "MontoNeto")
                Dim lblMontoNetoJunio As Label = CType(e.Row.FindControl("lblMontoNetoJunio"), Label)
                lblMontoNetoJunio.Text = Funciones.ExtraerValor(sSql + "06", "MontoNeto")
                Dim lblMontoNetoJulio As Label = CType(e.Row.FindControl("lblMontoNetoJulio"), Label)
                lblMontoNetoJulio.Text = Funciones.ExtraerValor(sSql + "07", "MontoNeto")
                Dim lblMontoNetoAgosto As Label = CType(e.Row.FindControl("lblMontoNetoAgosto"), Label)
                lblMontoNetoAgosto.Text = Funciones.ExtraerValor(sSql + "08", "MontoNeto")
                Dim lblMontoNetoSetiembre As Label = CType(e.Row.FindControl("lblMontoNetoSetiembre"), Label)
                lblMontoNetoSetiembre.Text = Funciones.ExtraerValor(sSql + "09", "MontoNeto")
                Dim lblMontoNetoOctubre As Label = CType(e.Row.FindControl("lblMontoNetoOctubre"), Label)
                lblMontoNetoOctubre.Text = Funciones.ExtraerValor(sSql + "10", "MontoNeto")
                Dim lblMontoNetoNoviembre As Label = CType(e.Row.FindControl("lblMontoNetoNoviembre"), Label)
                lblMontoNetoNoviembre.Text = Funciones.ExtraerValor(sSql + "11", "MontoNeto")
                Dim lblMontoNetoDiciembre As Label = CType(e.Row.FindControl("lblMontoNetoDiciembre"), Label)
                lblMontoNetoDiciembre.Text = Funciones.ExtraerValor(sSql + "12", "MontoNeto")
            End If
        End If
    End Sub

    Protected Sub gdvFlujo_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)
        gdvFlujo.EditIndex = e.NewEditIndex
        LlenaGridFlujo(0)
    End Sub

    Protected Sub gdvFlujo_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = gdvFlujo.Rows(e.RowIndex)
        Dim sCodVersion As String = "P"
        Dim sCodDetalle As String = gdvFlujo.DataKeys(e.RowIndex).Value.ToString()
        Dim hdfRubro As HiddenField = DirectCast(row.FindControl("hdfRubro"), HiddenField)
        Dim txtMontoNetoEnero As TextBox = DirectCast(row.FindControl("txtMontoNetoEnero"), TextBox)
        Dim bResultado As Boolean = True
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "01", sCodDetalle, hdfRubro.Value, txtMontoNetoEnero.Text)
        Dim txtMontoNetoFebrero As TextBox = DirectCast(row.FindControl("txtMontoNetoFebrero"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "02", sCodDetalle, hdfRubro.Value, txtMontoNetoFebrero.Text)
        Dim txtMontoNetoMarzo As TextBox = DirectCast(row.FindControl("txtMontoNetoMarzo"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "03", sCodDetalle, hdfRubro.Value, txtMontoNetoMarzo.Text)
        Dim txtMontoNetoAbril As TextBox = DirectCast(row.FindControl("txtMontoNetoAbril"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "04", sCodDetalle, hdfRubro.Value, txtMontoNetoAbril.Text)
        Dim txtMontoNetoMayo As TextBox = DirectCast(row.FindControl("txtMontoNetoMayo"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "05", sCodDetalle, hdfRubro.Value, txtMontoNetoMayo.Text)
        Dim txtMontoNetoJunio As TextBox = DirectCast(row.FindControl("txtMontoNetoJunio"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "06", sCodDetalle, hdfRubro.Value, txtMontoNetoJunio.Text)
        Dim txtMontoNetoJulio As TextBox = DirectCast(row.FindControl("txtMontoNetoJulio"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "07", sCodDetalle, hdfRubro.Value, txtMontoNetoJulio.Text)
        Dim txtMontoNetoAgosto As TextBox = DirectCast(row.FindControl("txtMontoNetoAgosto"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "08", sCodDetalle, hdfRubro.Value, txtMontoNetoAgosto.Text)
        Dim txtMontoNetoSetiembre As TextBox = DirectCast(row.FindControl("txtMontoNetoSetiembre"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "09", sCodDetalle, hdfRubro.Value, txtMontoNetoSetiembre.Text)
        Dim txtMontoNetoOctubre As TextBox = DirectCast(row.FindControl("txtMontoNetoOctubre"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "10", sCodDetalle, hdfRubro.Value, txtMontoNetoOctubre.Text)
        Dim txtMontoNetoNoviembre As TextBox = DirectCast(row.FindControl("txtMontoNetoNoviembre"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "11", sCodDetalle, hdfRubro.Value, txtMontoNetoNoviembre.Text)
        Dim txtMontoNetoDiciembre As TextBox = DirectCast(row.FindControl("txtMontoNetoDiciembre"), TextBox)
        bResultado = bResultado And Funciones.ActualizarDetalle(sCodVersion, cboPeriodo.Text + "12", sCodDetalle, hdfRubro.Value, txtMontoNetoDiciembre.Text)
        If (bResultado) Then
            lblMensaje.Text = "La información se ha grabado correctamente"
        Else
            lblMensaje.Text = "Se ha producido un error en uno de los meses"
        End If

        gdvFlujo.EditIndex = -1
        LlenaGridFlujo(0)
    End Sub

    Protected Sub gdvFlujo_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        gdvFlujo.EditIndex = -1
        LlenaGridFlujo(0)
    End Sub

End Class
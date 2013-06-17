Imports System.Data.OleDb
Imports System.Configuration

Public Class ValidacionFC
    Inherits System.Web.UI.Page

    Dim dMontoNeto As Decimal = 0
    Dim dIGV As Decimal = 0
    Dim dMontoBruto As Decimal = 0
    Dim dMontoBruto2 As Decimal = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim IdUsuario, CodTipo As String
        Dim Pagina As String = "ValidacionFC.aspx"
        Dim flag As Boolean = False

        If Not Session("IdUsuario") Is Nothing Then
            IdUsuario = Session("IdUsuario").ToString : CodTipo = Session("CodTipo").ToString
            flag = (CodTipo = "ADM" Or FuncionesSeguridad.ValidaAcceso(IdUsuario, Pagina))
        End If
        If Not flag Then Response.Redirect("Aviso.aspx")

        If Not IsPostBack Then
            CargaCombos()
            txtFechaIni.Text = Date.Now.ToString("yyyy-MM") & "-01"
            txtFechaFin.Text = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
            'gdvCuenta.PageSize = Funciones.iREGxPAG
            LlenaGridCuenta()
            LlenaGridCuenta2()
        End If
        'If hdfAccion.Value = "Recargar" Then
        '    LlenaGridCuenta()
        '    LlenaGridCuenta2()
        'End If
    End Sub

    Sub CargaCombos()
        cboCtaBanco.DataSource = LlenaCtasBancos()
        cboCtaBanco.DataTextField = "CuentaBanco"
        cboCtaBanco.DataValueField = "CodCuentaBanco"
        cboCtaBanco.DataBind()
    End Sub

    Sub LlenaGridCuenta()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()
        sql = "Select C.CodCuenta, C.Cuenta, sum(MontoBaseUSD) as MontoNeto, sum(MontoIGVUSD) as IGV, sum(MontoUSD) as MontoBruto,  "
        sql &= "case when Cuenta like 'ING-%' or Cuenta like 'OIN-%' then sum(MontoUSD) else -1 * sum(MontoUSD) end as MontoBruto2  "
        sql &= "From Movimiento M join CuentaFC C on M.CodCuentaFC = C.CodCuenta "
        sql &= "Where CodCuentaFC in (select distinct CodCtaOrigen from CtaDetalleRep where CodCtaOrigen is not null) "
        sql &= "and Fecha between '" & txtFechaIni.Text & "' and '" & txtFechaFin.Text & "' "
        If cboCtaBanco.SelectedIndex > 0 Then sql &= "and CodCuentaBanco = " & cboCtaBanco.SelectedValue & " "
        sql &= "Group by C.CodCuenta, C.Cuenta "
        sql &= "ORDER BY C.Cuenta"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCuenta.DataSource = dtaset.Tables(0)
        gdvCuenta.DataBind()

        cn.Close()
    End Sub

    Sub LlenaGridCuenta2()
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As DataSet
        Dim dtadap As OleDbDataAdapter

        cn = New OleDbConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()

        sql = "Select C.CodCuenta, C.Cuenta, sum(MontoBaseUSD) as MontoNeto, sum(MontoIGVUSD) as IGV, sum(MontoUSD) as MontoBruto,  "
        sql &= "case when Cuenta like 'ING-%' or Cuenta like 'OIN-%' then sum(MontoUSD) else -1 * sum(MontoUSD) end as MontoBruto2  "
        sql &= "From Movimiento M join CuentaFC C on M.CodCuentaFC = C.CodCuenta "
        sql &= "Where CodCuentaFC not in (select distinct CodCtaOrigen from CtaDetalleRep where CodCtaOrigen is not null) "
        sql &= "and Fecha between '" & txtFechaIni.Text & "' and '" & txtFechaFin.Text & "' "
        If cboCtaBanco.SelectedIndex > 0 Then sql &= "and CodCuentaBanco = " & cboCtaBanco.SelectedValue & " "
        sql &= "Group by C.CodCuenta, C.Cuenta "
        sql &= "ORDER BY C.Cuenta"

        cmd = New OleDbCommand(sql, cn)
        dtadap = New OleDbDataAdapter(cmd)
        dtaset = New DataSet()
        dtadap.Fill(dtaset)
        gdvCuenta2.DataSource = dtaset.Tables(0)
        gdvCuenta2.DataBind()

        cn.Close()
    End Sub

    Protected Sub gdvCuenta_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gdvCuenta.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            dMontoNeto += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoNeto"))
            dIGV += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "IGV"))
            dMontoBruto += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoBruto"))
            dMontoBruto2 += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoBruto2"))
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "Total: "
            e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(3).Text = dMontoNeto.ToString("c")
            e.Row.Cells(4).Text = dIGV.ToString("c")
            e.Row.Cells(5).Text = dMontoBruto.ToString("c")
            e.Row.Cells(6).Text = dMontoBruto2.ToString("c")
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(6).HorizontalAlign = HorizontalAlign.Right
            e.Row.Font.Bold = True
        End If
    End Sub

    Protected Sub gdvCuenta2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gdvCuenta2.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            dMontoNeto += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoNeto"))
            dIGV += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "IGV"))
            dMontoBruto += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoBruto"))
            dMontoBruto2 += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "MontoBruto2"))
        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "Total General: "
            e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(3).Text = dMontoNeto.ToString("c")
            e.Row.Cells(4).Text = dIGV.ToString("c")
            e.Row.Cells(5).Text = dMontoBruto.ToString("c")
            e.Row.Cells(6).Text = dMontoBruto2.ToString("c")
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(6).HorizontalAlign = HorizontalAlign.Right
            e.Row.Font.Bold = True
        End If
    End Sub

    Protected Sub btnDescargaExcel_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()
        gdvCuenta.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(gdvCuenta)
        frm.Controls.Add(gdvCuenta2)
        frm.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=FlujoCaja.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click
        LlenaGridCuenta()
        LlenaGridCuenta2()
    End Sub

    Protected Sub cboCtaBanco_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboCtaBanco.SelectedIndexChanged
        LlenaGridCuenta()
        LlenaGridCuenta2()
    End Sub

End Class
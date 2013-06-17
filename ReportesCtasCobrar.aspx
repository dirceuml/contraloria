<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesCtasCobrar.aspx.vb" Inherits="ATVContraloria.ReportesCtasCobrar" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Cuentas por Cobrar</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">

    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporteCtasPorCobrar" runat="server" Font-Size="Small">Generar Reporte Cuentas Por Cobrar (Contabilidad)</asp:LinkButton><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporteCtasPorCobrarContraloria" runat="server" Font-Size="Small">Generar Reporte Cuentas Por Cobrar (Contraloria)</asp:LinkButton><br/><br/>

    <asp:Label ID="lblFecha" runat="server" Text="A fecha:"></asp:Label>
    <asp:TextBox ID="txtFecha" runat="server" CssClass="InputText10" Width="60" MaxLength="10"></asp:TextBox>&nbsp;<img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFecha,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" />&nbsp;&nbsp;
    <asp:LinkButton ID="lnkGenerarReporteFacturasPorCobrar" runat="server" Font-Size="Small">Generar Reporte Facturas Por Cobrar (Contabilidad)</asp:LinkButton>

</asp:Content>

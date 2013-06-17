<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesEvolVentas.aspx.vb" Inherits="ATVContraloria.ReportesEvolVentas" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Evolución Ventas</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">

    <asp:Label ID="lblFecha" runat="server" Text="A fecha:"></asp:Label>
    <asp:TextBox ID="txtFecha" runat="server" CssClass="InputText10" Width="60" MaxLength="10"></asp:TextBox>&nbsp;<img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFecha,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" /><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporteEvolucionVentas" runat="server" Font-Size="Small">Generar Reporte Evolución Ventas</asp:LinkButton><br/><br/>

</asp:Content>

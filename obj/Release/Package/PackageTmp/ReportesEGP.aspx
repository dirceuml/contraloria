<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesEGP.aspx.vb" Inherits="ATVContraloria.ReportesEGP" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">EGP</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">
    <asp:Label ID="lblMensaje" runat="server" Text="" ForeColor="Red" Font-Bold="true" CssClass="fontTextRojo10"></asp:Label>
    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporte" runat="server" Font-Size="Small">Generar Reporte EGP</asp:LinkButton><br/><br/>

    Comparativo Cuentas
    <asp:DropDownList ID="cboCentroCosto" runat="server" CssClass="InputText10"></asp:DropDownList>
    <asp:DropDownList ID="cboRubroCosto" runat="server" CssClass="InputText10"></asp:DropDownList>&nbsp;
    <asp:LinkButton ID="lnkGenerarReporteComparativoCuentas" runat="server" Font-Size="Small">Generar Comparativo Cuentas</asp:LinkButton>

</asp:Content>

<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesRentabilidad.aspx.vb" Inherits="ATVContraloria.ReportesRentabilidad" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Rentabilidad</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">

    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>
    Distribución Costos Indirectos. Basado en Costos Directos (%) <asp:TextBox ID="txtPeso" runat="server" CssClass="InputText10" Width="30" MaxLength="5" AutoPostBack="true"></asp:TextBox>&nbsp;
    Basado en Horas (%) <asp:TextBox ID="txtPeso2" runat="server" CssClass="InputText10" Width="30" MaxLength="5" Enabled="false" ></asp:TextBox><br />
    <asp:LinkButton ID="lnkGenerarReporteRentabilidad" runat="server" Font-Size="Small">Generar Reporte Rentabilidad</asp:LinkButton>
</asp:Content>

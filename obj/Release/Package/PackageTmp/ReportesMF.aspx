<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesMF.aspx.vb" Inherits="ATVContraloria.ReportesMF" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Material Fílmico</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">

    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporteMaterialFilmico" runat="server" Font-Size="Small">Generar Reporte Material Fílmico</asp:LinkButton>

</asp:Content>

﻿<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesVentasGrupo.aspx.vb" Inherits="ATVContraloria.ReportesVentasGrupo" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Ventas Grupo</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">
    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>

    <asp:LinkButton ID="lnkGeneraReporteVentasGrupo" runat="server" Font-Size="Small">Generar Reporte Ventas Grupo</asp:LinkButton>

</asp:Content>

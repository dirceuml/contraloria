<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesConsumos.aspx.vb" Inherits="ATVContraloria.ReportesConsumos" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Consumos por Facturar</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">
    1) Seleccione Periodo: <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>
    2) Archivo de Consumos <b>Pagados</b> a Importar:<asp:FileUpload ID="fupConsumos" runat="server" />&nbsp;
    <asp:LinkButton ID="lnkImportar" runat="server" Font-Size="Small">Importar</asp:LinkButton><br/>
    <asp:Label ID="lblEstado" runat="server" Text="" ForeColor="Red"></asp:Label><br/><br/>
    3) (Opcional) Archivo de Consumos <b>Pagados (Clientes Inactivos)</b> a Importar:<asp:FileUpload ID="fupConsumosInactivos" runat="server" />&nbsp;
    <asp:LinkButton ID="lnkImportarInactivos" runat="server" Font-Size="Small">Importar</asp:LinkButton>&nbsp;&nbsp;&nbsp;
    <asp:LinkButton ID="lnkEjemploInactivos" runat="server" Font-Size="Small">Descargar Ejemplo</asp:LinkButton><br/>
    <asp:Label ID="lblEstadoInactivos" runat="server" Text="" ForeColor="Red"></asp:Label><br/><br/>
    4) <asp:LinkButton ID="lnkGenerarReporte" runat="server" Font-Size="Small">Generar Reporte Consumos</asp:LinkButton>
</asp:Content>

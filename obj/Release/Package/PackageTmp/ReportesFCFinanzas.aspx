<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ReportesFCFinanzas.aspx.vb" Inherits="ATVContraloria.ReportesFCFinanzas" %>
<asp:Content ID="cabecera" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="titulo" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Reportes : </span><span class="fontTitulo1b">Flujo Caja (Finanzas)</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="contenido" ContentPlaceHolderID="contenido" runat="server">

    <asp:RadioButtonList ID="rdbAmbito" runat="server" Font-Size="Small">
        <asp:ListItem Value="NAC" Selected="True">ATV Perú</asp:ListItem>
        <%--<asp:ListItem Value="EXT">Exterior</asp:ListItem>--%>
    </asp:RadioButtonList><br/>
    <asp:Label ID="lblPeriodo" runat="server" Text="Período:"></asp:Label>
    <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>

    <asp:LinkButton ID="lnkGenerarReporteTesoreria" runat="server" Font-Size="Small">Generar Reporte Flujo Caja Mensual (Tesoreria)</asp:LinkButton><br/><br/>

    <asp:Label ID="lblFechaIni" runat="server" Text="Fecha Inicio:"></asp:Label>
    <asp:TextBox ID="txtFechaIni" runat="server" CssClass="InputText10" Width="60" MaxLength="10"></asp:TextBox>&nbsp;<img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFechaIni,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" />&nbsp;&nbsp;
    <asp:Label ID="lblFechaFin" runat="server" Text="Fecha Fin:"></asp:Label>
    <asp:TextBox ID="txtFechaFin" runat="server" CssClass="InputText10" Width="60" MaxLength="10"></asp:TextBox>&nbsp;<img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFechaFin,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" /><br/><br/>
    <asp:LinkButton ID="lnkFCSemanal2" runat="server" Font-Size="Small">Generar Reporte Flujo Caja Semanal</asp:LinkButton><br/><br/>
    <asp:LinkButton ID="lnkFCSemanal2d" runat="server" Font-Size="Small">Generar Reporte Flujo Caja Diario</asp:LinkButton><br/><br/>
    <asp:LinkButton ID="lnkFCSemanal2b" runat="server" Font-Size="Small">Generar Reporte Flujo Caja RAG</asp:LinkButton><br/><br/>

</asp:Content>

<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Principal.aspx.vb" Inherits="ATVContraloria.Principal" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%">
            <span class="fontTitulo1">Bienvenido(a) : </span>
            <span class="fontTitulo1b"><asp:Label ID="lblUsuario" runat="server" Text="Label"></asp:Label></span>
      </td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    </td>
</tr>
</table>
</asp:Content>

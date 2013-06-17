<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Personas.aspx.vb" Inherits="ATVContraloria.Personas" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Personas</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>

</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
    <asp:HiddenField ID="hdfAccion" Value="" runat="server" />
    <asp:HiddenField ID="hdfNroRegistros" Value="0" runat="server" />
    <table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
    <tr><td style="background-color:#4d81b3;"><font class="fontTextBlanco"><b>Filtros</b></font></td></tr>
    <tr class="TDList1">
        <td>
            <font class="fontTextNormal">Persona: </font>&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="txtPersona" runat="server" CssClass="InputText10" Width="100" MaxLength="10"></asp:TextBox>&nbsp;&nbsp;&nbsp;
            <asp:Button ID="btnConsultar" runat="server" Text="Consultar" OnClick="btnConsultar_Click" Height="20px" CssClass ="fontBoton" />
        </td>
    </tr>
    <tr><td style="width:1px; height:2px;"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr><td style="background-color:#4d81b3;"><font class="txtBlanco"><b>
        <asp:Label ID="lblNumRegistros" runat="server" Text=""></asp:Label></b></font></td>
    </tr>
    <tr><td style="width:1px; height:2px;"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr>
        <td align="center" >
        <asp:GridView ID="gdvPersona" runat="server" AutoGenerateColumns="False" Width="100%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowPaging="True" 
            GridLines="Both" onpageindexchanging="gdvPersona_PageIndexChanging" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Código" DataField="CodPersona" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50"></asp:BoundField>
                <asp:BoundField HeaderText="Persona" DataField="Persona" ItemStyle-HorizontalAlign="left"></asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        </td>
    </tr>
    </table>

</asp:Content>
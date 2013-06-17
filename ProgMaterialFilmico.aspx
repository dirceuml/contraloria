<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ProgMaterialFilmico.aspx.vb" Inherits="ATVContraloria.ProgMaterialFilmico" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Programación Material Fílmico: <asp:Label ID="lblMaterial" runat="server" Text=""></asp:Label></span></td>
      <td width="39%" align="right" nowrap="nowrap"><asp:Button ID="btnExcel" runat="server" Text="Exportar a Excel" onclick="btnExcel_Click" Height="20px" Width="120px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" /></td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<asp:ScriptManager ID="ScriptManager1" runat="server">
</asp:ScriptManager>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
<ContentTemplate>

<table border="0" cellspacing="0" cellpadding="5" width="100%" align="center">
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:440px;text-align:center;overflow:auto;">
    <asp:GridView ID="gdvConsumoMaterialFilmico" runat="server" AutoGenerateColumns="False" Width="360"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3"
        GridLines="Both" DataKeyNames="IdConsumoMaterial">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>            
            <asp:BoundField HeaderText="Periodo" DataField="Periodo" ItemStyle-Width="100" ></asp:BoundField>
            <asp:BoundField HeaderText="Capitulos Prog." DataField="CantCapitulosProg" ReadOnly="true" DataFormatString="{0:n0}" ItemStyle-Width="80" ItemStyle-HorizontalAlign="Right" ></asp:BoundField>
            <asp:BoundField HeaderText="Importe US$" DataField="MontoUSD" ReadOnly="true" DataFormatString="{0:n0}" ItemStyle-Width="80" ItemStyle-HorizontalAlign="Right" ></asp:BoundField>
        </Columns>
        <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#DCDCDC" />
    </asp:GridView><br/>
    <asp:GridView ID="gdvProgMaterialFilmico" runat="server" AutoGenerateColumns="False" Width="480"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3"
        GridLines="Both" DataKeyNames="IdProgMaterial">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>            
            <asp:BoundField HeaderText="Periodo" DataField="Periodo" ItemStyle-Width="100" ></asp:BoundField>
            <asp:BoundField HeaderText="Fecha" DataField="FechaProg" DataFormatString="{0:dd/MM/yyyy}" ItemStyle-Width="120" ></asp:BoundField>
            <asp:BoundField HeaderText="Programa" DataField="Programa" ItemStyle-Width="120" ></asp:BoundField>
            <asp:BoundField HeaderText="Capítulo" DataField="NumCapitulo" ItemStyle-Width="60" ></asp:BoundField>
            <asp:BoundField HeaderText="Rating" DataField="Rating" ItemStyle-Width="60"></asp:BoundField>
        </Columns>
        <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#DCDCDC" />
    </asp:GridView>
    <asp:GridView ID="gdvConsumoMaterialFilmicoExp" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" EnableModelValidation="True" ForeColor="#333333" Width="180"
            GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
            <asp:BoundField HeaderText="Periodo" DataField="Periodo" ItemStyle-Width="100" ></asp:BoundField>
            <asp:BoundField HeaderText="Capitulos Prog." DataField="CantCapitulosProg" DataFormatString="{0:n0}" ItemStyle-Width="80" ></asp:BoundField>
            <asp:BoundField HeaderText="Importe US$" DataField="MontoUSD"  DataFormatString="{0:n0}" ItemStyle-Width="80" ></asp:BoundField>
        </Columns>
        <EditRowStyle BackColor="#7C6F57" />
        <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#E3EAEB" />
        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
    </asp:GridView>
    <asp:GridView ID="gdvProgMaterialFilmicoExp" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" EnableModelValidation="True" ForeColor="#333333" Width="480"
            GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
            <asp:BoundField HeaderText="Periodo" DataField="Periodo" ItemStyle-Width="100" ></asp:BoundField>
            <asp:BoundField HeaderText="Fecha" DataField="FechaProg" DataFormatString="{0:yyyy-MM-dd}" ItemStyle-Width="120" ></asp:BoundField>
            <asp:BoundField HeaderText="Programa" DataField="Programa" ItemStyle-Width="120" ></asp:BoundField>
            <asp:BoundField HeaderText="Capítulo" DataField="NumCapitulo" ItemStyle-Width="60" ></asp:BoundField>
            <asp:BoundField HeaderText="Rating" DataField="Rating" ItemStyle-Width="60" ></asp:BoundField>
        </Columns>
        <EditRowStyle BackColor="#7C6F57" />
        <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#E3EAEB" />
        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
    </asp:GridView>
    </div> 

    </td>
</tr>
</table>

</ContentTemplate>
</asp:UpdatePanel>

</asp:Content>

<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Areas.aspx.vb" Inherits="ATVContraloria.Areas" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Areas</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
    <asp:GridView ID="gdvArea" runat="server" AutoGenerateColumns="False" Width="98%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowSorting="True" onsorting="gdvArea_Sorting" 
        GridLines="Both" DataKeyNames="CodArea" 
        onrowdatabound="gdvArea_RowDataBound"
        onrowediting="gdvArea_RowEditing" 
        onrowcancelingedit="gdvArea_RowCancelingEdit" 
        onrowupdating="gdvArea_RowUpdating" >                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:BoundField HeaderText="Código" DataField="CodCentroCosto" ReadOnly="true" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodCentroCosto" ></asp:BoundField>
            <asp:BoundField HeaderText="Centro de Costo" DataField="CentroCosto" ReadOnly="true" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" SortExpression="CentroCosto" ></asp:BoundField>
            <asp:BoundField HeaderText="Código" DataField="CodArea" ReadOnly="true" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodArea" ></asp:BoundField>
            <asp:BoundField HeaderText="Area" DataField="Area" ReadOnly="true" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" SortExpression="Area" ></asp:BoundField>
            <asp:BoundField HeaderText="CodEstado" DataField="CodEstado" ReadOnly="true" ItemStyle-HorizontalAlign="left" Visible="false" ControlStyle-CssClass="fontText10" SortExpression="CodEstado" ></asp:BoundField>
            <asp:TemplateField HeaderText="Programa" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:Label ID="lnkPrograma" runat="server" Text="lnkPrograma" CssClass="fontText10" ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboPrograma" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="S" Text="SI"></asp:ListItem>
                    <asp:ListItem Value="N" Text="NO"></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" CausesValidation="False" ><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:LinkButton ID="lnkUpdate" runat="server" CausesValidation="True" CommandName="Update"><asp:Image ID="imgUpdate" runat="server" ImageUrl="images/save.gif" AlternateText="Grabar" /></asp:LinkButton>&nbsp;
                        &nbsp;
                        <asp:LinkButton ID="lnkCancel" runat="server" CausesValidation="False" CommandName="Cancel"><asp:Image ID="imgCancel" runat="server" ImageUrl="images/cancel.gif" AlternateText="Cancelar" /></asp:LinkButton>
                </EditItemTemplate>
            </asp:TemplateField>
        </Columns>
        <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#DCDCDC" />
    </asp:GridView>
    <br />
    </div> 
    </td>
</tr>
</table>
</asp:Content>

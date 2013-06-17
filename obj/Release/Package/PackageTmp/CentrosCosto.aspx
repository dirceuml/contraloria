<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="CentrosCosto.aspx.vb" Inherits="ATVContraloria.CentrosCosto" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Centros Costo</span></td>
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
    <asp:HiddenField ID="hdfOrden" runat="server" />
    <asp:HiddenField ID="hdfTipoOrden" runat="server" />
    <asp:GridView ID="gdvCentrosCosto" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowSorting="True" 
        GridLines="Both" DataKeyNames="CodCentroCosto" AllowPaging="True">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>            
<%--            <asp:TemplateField HeaderText="No." ItemStyle-HorizontalAlign="Right" HeaderStyle-Width="40" >
                <ItemTemplate>
                    <%# Container.DataItemIndex + 1 %>&nbsp;
                </ItemTemplate>
            </asp:TemplateField>--%>
            <asp:BoundField HeaderText="Código" DataField="CodCentroCosto" ReadOnly="true" SortExpression="CodCentroCosto" ItemStyle-Width="40"></asp:BoundField>
            <asp:BoundField HeaderText="Centro Costo" DataField="CentroCosto" ReadOnly="true" SortExpression="CentroCosto" ></asp:BoundField>
            <asp:TemplateField HeaderText="Tipo" SortExpression="Tipo" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="60" >
                <ItemTemplate>
                    <asp:Label ID="lblTipo" runat="server" Text='<%# Eval("Tipo") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboTipo" runat="server" Width="50" SelectedValue='<%# Eval("CodTipo") %>' Font-Size="Smaller" >
                        <asp:ListItem Value="D" Text="Directo" ></asp:ListItem> 
                        <asp:ListItem Value="I" Text="Indirecto" ></asp:ListItem>
                        <asp:ListItem Value="O" Text="Otro" ></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Deporte?" SortExpression="FlagDeporte" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="60" >
                <ItemTemplate>
                    <asp:Label ID="lblFlagDeporte" runat="server" Text='<%# Eval("FlagDeporte") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:CheckBox ID="chkFlagDeporte" runat="server" CssClass="InputText10" Checked='<%# If(Eval("FlagDeporte").ToString() = "S", True, False) %>' />
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Centro EGP" SortExpression="CentroCostoEGP" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="200" >
                <ItemTemplate>
                    <asp:Label ID="lblCentroCostoEGP" runat="server" Text='<%# Eval("CentroCostoEGP") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:TextBox ID="txtCentroCostoEGP" runat="server" Width="160" Text='<%# Eval("CentroCostoEGP") %>' Font-Size="Smaller" ></asp:TextBox>
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Grupo EGP" SortExpression="GrupoEGP" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="200" >
                <ItemTemplate>
                    <asp:Label ID="lblGrupoEGP" runat="server" Text='<%# Eval("GrupoEGP") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:TextBox ID="txtCodGrupoEGP" runat="server" Width="20" Text='<%# Eval("CodGrupoEGP") %>' Font-Size="Smaller" ></asp:TextBox>
                    <asp:TextBox ID="txtGrupoEGP" runat="server" Width="160" Text='<%# Eval("GrupoEGP") %>' Font-Size="Smaller" ></asp:TextBox>
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="60" >
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
    <asp:GridView ID="gdvCentrosCostoExp" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" EnableModelValidation="True" ForeColor="#333333" 
            GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
<%--            <asp:TemplateField HeaderText="No." HeaderStyle-Width="40" ItemStyle-HorizontalAlign="Right" >
                <ItemTemplate>
                    <%# Container.DataItemIndex + 1 %>
                </ItemTemplate>
            </asp:TemplateField>--%>
            <asp:BoundField HeaderText="Código" DataField="CodCentroCosto" HeaderStyle-Width="50" />
            <asp:BoundField HeaderText="Centro Costo" DataField="CentroCosto" HeaderStyle-Width="260" />
            <asp:BoundField HeaderText="Tipo" DataField="Tipo" HeaderStyle-Width="80" ItemStyle-HorizontalAlign="Center" />
            <asp:BoundField HeaderText="Deporte?" DataField="FlagDeporte" HeaderStyle-Width="80" ItemStyle-HorizontalAlign="Center" />
            <asp:BoundField HeaderText="Centro EGP" DataField="CentroCostoEGP" HeaderStyle-Width="260" />
            <asp:BoundField HeaderText="Grupo EGP" DataField="GrupoEGP" HeaderStyle-Width="260" />
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

<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Accesos.aspx.vb" Inherits="ATVContraloria.Accesos" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimientos : </span><span class="fontTitulo1b">Accesos</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:500px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" ForeColor="Red" Font-Bold="true" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvAcceso" runat="server" AutoGenerateColumns="False" Width="98%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdAcceso" >                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Sección" ItemStyle-Width="120px" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate>&nbsp;&nbsp;<%# Eval("Seccion")%></ItemTemplate>
                <EditItemTemplate>
                    &nbsp;&nbsp;<asp:DropDownList ID="cboSeccion" runat="server" DataValueField="Seccion" DataTextField="Seccion" CssClass="fontText10" ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    &nbsp;&nbsp;<asp:DropDownList ID="cboSeccion" runat="server" DataValueField="Seccion" DataTextField="Seccion" CssClass="fontText10" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Orden"  ItemStyle-Width="40px" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("Orden")%></ItemTemplate>
                <EditItemTemplate>
                   <asp:TextBox ID="txtOrden" runat="server" Text='<%# Bind("Orden")%>' Width="20px" CssClass="fontText10" ></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtOrden" runat="server" Text='' CssClass="fontText10" Width="20px" ></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Acceso" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate>&nbsp;<%# Eval("Acceso")%></ItemTemplate>
                <EditItemTemplate>
                    &nbsp;<asp:TextBox ID="txtAcceso" runat="server" Text='<%# Bind("Acceso")%>' Width="200px" CssClass="fontText10" ></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    &nbsp;<asp:TextBox ID="txtAcceso" runat="server" Text='' CssClass="fontText10" Width="200px" ></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Página" ItemStyle-Width="180px" ItemStyle-HorizontalAlign="Center" FooterStyle-HorizontalAlign="Left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate>&nbsp;<%# Eval("Pagina")%></ItemTemplate>
                <EditItemTemplate>
                    &nbsp;<asp:TextBox ID="txtPagina" runat="server" Text='<%# Bind("Pagina")%>' Width="160px" CssClass="fontText10" ></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    &nbsp;<asp:TextBox ID="txtPagina" runat="server" Text='' Width="160px" CssClass="fontText10" ></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Acción" ItemStyle-Width="60px" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" CausesValidation="False" ><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                    &nbsp;
                    <span onclick="return confirm('Esta usted seguro en borrar el registro?')"><asp:LinkButton ID="lnkDelete" runat="server" CommandName="Delete"><asp:Image ID="ImgDelete" runat="server" ImageUrl="images/delete.gif" AlternateText="Eliminar" /></asp:LinkButton></span>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:LinkButton ID="lnkUpdate" runat="server" CausesValidation="True" CommandName="Update"><asp:Image ID="imgUpdate" runat="server" ImageUrl="images/save.gif" AlternateText="Grabar" /></asp:LinkButton>&nbsp;
                    &nbsp;
                    <asp:LinkButton ID="lnkCancel" runat="server" CausesValidation="False" CommandName="Cancel"><asp:Image ID="imgCancel" runat="server" ImageUrl="images/cancel.gif" AlternateText="Cancelar" /></asp:LinkButton>
                </EditItemTemplate>
                <FooterTemplate><asp:LinkButton ID="lnkInsert" runat="server" OnClick="lnkInsert_Click" CausesValidation="False" ><asp:Image ID="imgInsert" runat="server" ImageUrl="images/new.gif" AlternateText="Crear uno nuevo" /></asp:LinkButton></FooterTemplate>
            </asp:TemplateField>
        </Columns>
        <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#DCDCDC" />
    </asp:GridView>
    </div> 
    </td>
</tr>
</table>
</asp:Content>

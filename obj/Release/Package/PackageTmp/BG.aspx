﻿<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="BG.aspx.vb" Inherits="ATVContraloria.BG" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimiento : </span><span class="fontTitulo1b">Balance General</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    <asp:Button ID="btnDescargaExcel" runat="server" Text="Descargar a Excel" Height="20px" Width="120px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
    <div style="background-color:#FFFFFF;height:500px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvCtaBG" runat="server" AutoGenerateColumns="False" Width="98%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdCtaBG" >                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Detalle" HeaderStyle-Width="50">
                <ItemTemplate><asp:HyperLink ID="hlnkDetalle" runat="server" ImageUrl="images/lupa.png" ToolTip="Ver detalle" Target="_self"></asp:HyperLink></ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Orden" HeaderStyle-Width="50" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("Orden")%></ItemTemplate>
                <EditItemTemplate>
                    <asp:TextBox ID="txtOrden" runat="server" Text='<%# Bind("Orden")%>' Width="30px" CssClass="fontText10" ></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtOrden" runat="server" Text='' CssClass="fontText10" Width="30px" ></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cuenta Balance General" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("CtaBG")%></ItemTemplate>
                <EditItemTemplate>
                    <asp:TextBox ID="txtCtaBG" runat="server" Text='<%# Bind("CtaCtaBG")%>' Width="200px" CssClass="fontText10" ></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtCtaBG" runat="server" Text='' CssClass="fontText10" Width="200px" ></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Sección" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate><%# Eval("CodSeccion")%></ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboCodSeccion" runat="server" CssClass="fontText10" DataValueField="CodSeccion" DataTextField="CodSeccion" ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                   <asp:DropDownList ID="cboCodSeccion" runat="server" CssClass="fontText10" DataValueField="CodSeccion" DataTextField="CodSeccion" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Modif?" HeaderStyle-Width="50" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:Label ID="lblFlagModif" runat="server" Text="" CssClass="fontText10" ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboFlagModif" runat="server" CssClass="fontText10" >
                        <asp:ListItem Value="S" Text="Sí"></asp:ListItem>
                        <asp:ListItem Value="N" Text="No"></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboFlagModif" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="S" Text="Sí"></asp:ListItem>
                    <asp:ListItem Value="N" Text="No"></asp:ListItem>
                    </asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Acción" ItemStyle-HorizontalAlign="Center">
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
    <asp:GridView ID="gdvCtaBGExp" runat="server" AutoGenerateColumns="False" Visible="False">
        <Columns>
            <asp:BoundField HeaderText="Orden" DataField="Orden" ItemStyle-Width="50px"/>
            <asp:BoundField HeaderText="Cuenta Balance General" DataField="CtaBG" ItemStyle-Width="300px" />
            <asp:BoundField HeaderText="Sección" DataField="CodSeccion" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="60px"/>
            <asp:BoundField HeaderText="Cod. Cuenta" DataField="CodCuenta" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="100px" />
            <asp:BoundField HeaderText="Cuenta Contable SISMED" DataField="Cuenta" ItemStyle-Width="520px" />
        </Columns>
    </asp:GridView>
    </div> 
    </td>
</tr>
</table>
</asp:Content>

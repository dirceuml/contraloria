﻿<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ERDetalle.aspx.vb" Inherits="ATVContraloria.ERDetalle" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimientos : </span><span class="fontTitulo1b"><a href="ER.aspx" style="text-decoration:underline;">Estado Resultados</a> -> Detalle</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<asp:HiddenField ID="hdfIdCtaER" Value="" runat="server" />
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr><td><span class="fontText12"><b>&nbsp;&nbsp;&nbsp;<asp:Label ID="lblCtaER" runat="server" Text="lblCtaER"></asp:Label></b></span><br /><br /></td></tr>
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:500px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvCtaERDet" runat="server" AutoGenerateColumns="False" Width="98%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdCtaERDet" >                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Signo" ItemStyle-Width="60px" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate><%# IIf(Eval("Signo") = "1", "+", "-")%></ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboSigno" runat="server" CssClass="fontText10" >
                        <asp:ListItem Value="1" Text=" + "></asp:ListItem>
                        <asp:ListItem Value="-1" Text=" - "></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboSigno" runat="server" CssClass="fontText10" >
                        <asp:ListItem Value="1" Text=" + "></asp:ListItem>
                        <asp:ListItem Value="-1" Text=" - "></asp:ListItem>
                    </asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cuenta Contable SISMED" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Left" >
                <ItemTemplate>&nbsp;<%# Eval("Cuenta")%></ItemTemplate>
                <EditItemTemplate>
                    &nbsp;<asp:DropDownList ID="cboCodCtaOrigen" runat="server" CssClass="fontText10" DataValueField="CodCuenta" DataTextField="Cuenta" ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    &nbsp;<asp:DropDownList ID="cboCodCtaOrigen" runat="server" CssClass="fontText10" DataValueField="CodCuenta" DataTextField="Cuenta" ></asp:DropDownList>
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
    <br />
    </div> 
    </td>
</tr>
</table>
</asp:Content>

﻿<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="FlujoCajaNDetalle.aspx.vb" Inherits="ATVContraloria.FlujoCajaNDetalle" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimientos : </span><span class="fontTitulo1b"><a href="FlujoCajaN.aspx" style="text-decoration:underline;">Flujo de Caja (Nuevo Formato)</a> -> Detalle</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<asp:HiddenField ID="hdfIdCtaFlujoNRep" Value="" runat="server" />
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr><td><span class="fontText12"><b>&nbsp;&nbsp;&nbsp;Cuenta: <asp:Label ID="lblCtaFlujoNRep" runat="server" Text="lblCtaFlujoNRep"></asp:Label></b></span></td></tr>
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:500px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvCtaDetalleNRep" runat="server" AutoGenerateColumns="False" Width="98%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdCtaDetalleNRep" 
        onrowdatabound="gdvCtaDetalleNRep_RowDataBound"
        onrowediting="gdvCtaDetalleNRep_RowEditing" 
        OnRowDeleting="gdvCtaDetalleNRep_RowDeleting"
        onrowcancelingedit="gdvCtaDetalleNRep_RowCancelingEdit" 
        onrowupdating="gdvCtaDetalleNRep_RowUpdating" >                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Orden" HeaderStyle-Width="50" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("Orden")%></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtOrden" runat="server" Text='<%# Bind("Orden")%>' Width="50px" CssClass="fontText10" ></asp:TextBox></EditItemTemplate>
                <FooterTemplate><asp:TextBox ID="txtOrdenNew" runat="server" Text='' CssClass="fontText10" ></asp:TextBox></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cuenta Detalle" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("CtaDetalleNRep")%></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtCtaDetalleNRep" runat="server" Text='<%# Bind("CtaDetalleNRep")%>' Width="200px" CssClass="fontText10" ></asp:TextBox></EditItemTemplate>
                <FooterTemplate><asp:TextBox ID="txtCtaDetalleNRepNew" runat="server" Text='' CssClass="fontText10" ></asp:TextBox></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Signo" HeaderStyle-Width="50" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:Label ID="lnkSigno" runat="server" Text="lnkSigno" CssClass="fontText10" ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboSigno" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="0" Text="   "></asp:ListItem>
                    <asp:ListItem Value="1" Text=" + "></asp:ListItem>
                    <asp:ListItem Value="-1" Text=" - "></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboSignoNew" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="0" Text="   "></asp:ListItem>
                    <asp:ListItem Value="1" Text=" + "></asp:ListItem>
                    <asp:ListItem Value="-1" Text=" - "></asp:ListItem>
                    </asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cta Origen" ItemStyle-HorizontalAlign="left">
                <ItemTemplate><asp:Label ID="lnkCodCtaOrigen" runat="server" Text="lnkCodCtaOrigen" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:DropDownList ID="cboCodCtaOrigen" runat="server" CssClass="fontText10" ></asp:DropDownList></EditItemTemplate>
                <FooterTemplate><asp:DropDownList ID="cboCodCtaOrigenNew" runat="server" CssClass="fontText10" ></asp:DropDownList></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Tipo Detalle" ItemStyle-HorizontalAlign="left">
                <ItemTemplate><asp:Label ID="lnkTipoDetalle" runat="server" Text="lnkTipoDetalle" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboTipoDetalle" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="" Text=" "></asp:ListItem>
                    <asp:ListItem Value="ARE" Text="Area"></asp:ListItem>
                    <asp:ListItem Value="GRP" Text="Grupo"></asp:ListItem>
                    <asp:ListItem Value="PER" Text="Persona"></asp:ListItem>
                    <asp:ListItem Value="PRG" Text="Programa"></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboTipoDetalleNew" runat="server" CssClass="fontText10" >
                    <asp:ListItem Value="" Text=" "></asp:ListItem>
                    <asp:ListItem Value="ARE" Text="Area"></asp:ListItem>
                    <asp:ListItem Value="GRP" Text="Grupo"></asp:ListItem>
                    <asp:ListItem Value="PER" Text="Persona"></asp:ListItem>
                    <asp:ListItem Value="PRG" Text="Programa"></asp:ListItem>
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
                <FooterTemplate><asp:LinkButton ID="lnkInsert" runat="server" OnClick="btnNuevo_Click" CausesValidation="False" ><asp:Image ID="imgInsert" runat="server" ImageUrl="images/new.gif" AlternateText="Crear uno nuevo" /></asp:LinkButton></FooterTemplate>
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
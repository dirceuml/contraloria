﻿<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="CuentasFC.aspx.vb" Inherits="ATVContraloria.CuentasFC" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimientos : </span><span class="fontTitulo1b">Cuentas Flujo Caja</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<AjaxToolKit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></AjaxToolKit:ToolkitScriptManager>
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td style="background-color:#4d81b3;">&nbsp;&nbsp;<font class="fontTextBlanco"><b>Cuentas de Flujo Caja</b></font></td>
    <td><img src="images/trans.gif" alt="" width="30" height="1" /></td>
    <td style="background-color:#4d81b3;">&nbsp;&nbsp;<font class="fontTextBlanco"><b>Cuentas Nuevas</b></font></td>
</tr>
<tr>
    <td>
    <asp:UpdatePanel ID="upnlCuentaFC" runat="server">
      <ContentTemplate>
        <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
        <asp:GridView ID="gdvCuentaFC" runat="server" AutoGenerateColumns="False" Width="97%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowSorting="True" onsorting="gdvCuentaFC_Sorting" 
            GridLines="Both" DataKeyNames="CodCuenta" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Código" DataField="CodCuenta" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodCuenta" ></asp:BoundField>
                <asp:BoundField HeaderText="Cuenta" DataField="Cuenta" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" SortExpression="Cuenta"></asp:BoundField>
                <asp:BoundField HeaderText="Estado" DataField="CodEstado" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodEstado"></asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        <br />
        </div>
      </ContentTemplate>
    </asp:UpdatePanel>
    </td>
    <td><img src="images/trans.gif" alt="" width="30" height="1" /></td>
    <td>
    <asp:UpdatePanel ID="upnlCuentaFC2" runat="server">
      <ContentTemplate>
        <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
        <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
        <asp:GridView ID="gdvCuentaFC2" runat="server" AutoGenerateColumns="False" Width="97%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true" AllowSorting="True" onsorting="gdvCuentaFC2_Sorting" 
            GridLines="Both" DataKeyNames="CodCuenta2" 
            onrowdatabound="gdvCuentaFC2_RowDataBound"
            onrowediting="gdvCuentaFC2_RowEditing" 
            OnRowDeleting="gdvCuentaFC2_RowDeleting"
            onrowcancelingedit="gdvCuentaFC2_RowCancelingEdit" 
            onrowupdating="gdvCuentaFC2_RowUpdating">                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Código" DataField="CodCuenta2" ReadOnly="true" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodCuenta2" ></asp:BoundField>
                <asp:TemplateField HeaderText="Cuenta" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" SortExpression="Cuenta2" >
                    <ItemTemplate><%# Eval("Cuenta2")%></ItemTemplate>
                    <EditItemTemplate><asp:TextBox ID="txtCuenta2" runat="server" Text='<%# Bind("Cuenta2")%>' Width="200px" CssClass="fontText10" ></asp:TextBox></EditItemTemplate>
                    <FooterTemplate><asp:TextBox ID="txtCuenta2New" runat="server" Text='' Width="200px" CssClass="fontText10" ></asp:TextBox></FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Estado" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ControlStyle-CssClass="fontText10" SortExpression="CodEstado" >
                    <ItemTemplate><asp:Label ID="lblCodEstado" runat="server" Text="lblCodEstado" CssClass="fontText10" ></asp:Label></ItemTemplate>
                    <EditItemTemplate><asp:DropDownList ID="cboCodEstado" runat="server" CssClass="fontText10" ></asp:DropDownList></EditItemTemplate>
                    <FooterTemplate><asp:DropDownList ID="cboCodEstadoNew" runat="server" CssClass="fontText10" ></asp:DropDownList></FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center">
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
      </ContentTemplate>
    </asp:UpdatePanel>
    </td>
</tr>
</table>
</asp:Content>
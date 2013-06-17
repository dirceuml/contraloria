<%@ Page Title="" Language="vb" AutoEventWireup="false" EnableEventValidation="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="PersonasFCS.aspx.vb" Inherits="ATVContraloria.PersonasFCS" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Actualización : </span><span class="fontTitulo1b">Personas en el Flujo de Caja</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">

<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
<td>
    <asp:Button ID="btnDescargaExcel" runat="server" Text="Excel" onclick="btnDescargaExcel_Click" Height="20px" Width="80px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
</td>
</tr>
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvPersonaS" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdPersonaS" 
        onrowdatabound="gdvPersonaS_RowDataBound"
        onrowediting="gdvPersonaS_RowEditing" 
        onRowDeleting="gdvPersonaS_RowDeleting"
        onrowcancelingedit="gdvPersonaS_RowCancelingEdit" 
        onrowupdating="gdvPersonaS_RowUpdating">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Grupo" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblGrupo" runat="server" Text="lblGrupo" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:DropDownList ID="cboCodGrupo" runat="server" CssClass="fontText10" ></asp:DropDownList></EditItemTemplate>
                <FooterTemplate><asp:DropDownList ID="cboCodGrupoNew" runat="server" CssClass="fontText10" ></asp:DropDownList></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Personas en el FC" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("PersonaS")%></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtPersonaS" runat="server" Text='<%# Bind("PersonaS")%>' CssClass="fontText10" ></asp:TextBox></EditItemTemplate>
                <FooterTemplate><asp:TextBox ID="txtPersonaSNew" runat="server" Text='' CssClass="fontText10" ></asp:TextBox></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Persona" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblPersona" runat="server" Text="lblPersona" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:DropDownList ID="cboCodPersona" runat="server" CssClass="fontText10" ></asp:DropDownList></EditItemTemplate>
                <FooterTemplate><asp:DropDownList ID="cboCodPersonaNew" runat="server" CssClass="fontText10" ></asp:DropDownList></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Acción" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" CausesValidation="False"><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                    &nbsp;
                    <span onclick="return confirm('Esta usted seguro en borrar el registro?')"><asp:LinkButton ID="lnkDelete" runat="server" CommandName="Delete"><asp:Image ID="ImgDelete" runat="server" ImageUrl="images/delete.gif" AlternateText="Eliminar" /></asp:LinkButton></span>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:LinkButton ID="lnkUpdate" runat="server" CausesValidation="True" CommandName="Update"><asp:Image ID="imgUpdate" runat="server" ImageUrl="images/save.gif" AlternateText="Grabar" /></asp:LinkButton>&nbsp;
                    <asp:LinkButton ID="lnkCancel" runat="server" CausesValidation="False" CommandName="Cancel"><asp:Image ID="imgCancel" runat="server" ImageUrl="images/cancel.gif" AlternateText="Cancelar" /></asp:LinkButton>
                </EditItemTemplate>
                <FooterTemplate><asp:LinkButton ID="lnkInsert" runat="server" OnClick="btnNuevo_Click" CausesValidation="False"><asp:Image ID="imgInsert" runat="server" ImageUrl="images/new.gif" AlternateText="Crear uno nuevo" /></asp:LinkButton></FooterTemplate>
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

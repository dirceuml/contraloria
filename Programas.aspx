<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Programas.aspx.vb" Inherits="ATVContraloria.Programas" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Programas</span></td>
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
<tr height="20">
    <td style="background-color:#4d81b3;">
        <font class="fontTextBlanco"><b>&nbsp;Programa</b></font>
        <asp:TextBox ID="txtProgramaB" runat="server" CssClass="InputText10" Width="100"></asp:TextBox>&nbsp;
        <asp:Button ID="btnConsultar" runat="server" Text="Consultar" Width="120px" Height="20px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
    </td>
</tr>
<tr>
    <td height="5"></td>
</tr>
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:440px;text-align:center;overflow:auto;">
    <asp:HiddenField ID="hdfOrden" runat="server" />
    <asp:HiddenField ID="hdfTipoOrden" runat="server" />
    <asp:GridView ID="gdvProgramas" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowSorting="True" 
        GridLines="Both" DataKeyNames="IdPrograma" AllowPaging="True">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>            
            <asp:TemplateField HeaderText="No." ItemStyle-HorizontalAlign="Right" HeaderStyle-Width="40" >
                <ItemTemplate>
                    <%# Container.DataItemIndex + 1 %>&nbsp;
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField HeaderText="Programa" DataField="Programa" ReadOnly="true" SortExpression="Programa" ></asp:BoundField>
            <asp:BoundField HeaderText="Género" DataField="Genero" ReadOnly="true" SortExpression="Genero" ItemStyle-Width="200" ></asp:BoundField>
            <asp:TemplateField HeaderText="M.Fílmico?" SortExpression="FlagMaterial" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="100" >
                <ItemTemplate>
                    <asp:Label ID="lblFlagMaterial" runat="server" Text='<%# Eval("FlagMaterial") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:CheckBox ID="chkFlagMaterial" runat="server" CssClass="InputText10" Checked='<%# If(Eval("FlagMaterial").ToString() = "S", True, False) %>' />
                </EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Grupo Programa" SortExpression="GrupoPrograma" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="240" >
                <ItemTemplate>
                    <asp:Label ID="lblGrupoPrograma" runat="server" Text='<%# Eval("GrupoPrograma") %>' ></asp:Label>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboGruposPrograma" runat="server" Width="220" Font-Size="Smaller"></asp:DropDownList><br/>
                    <asp:TextBox ID="txtGrupoPrograma" runat="server" Width="220" Font-Size="Smaller" ></asp:TextBox>
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
    <asp:GridView ID="gdvProgramasExp" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" EnableModelValidation="True" ForeColor="#333333" 
            GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
            <asp:TemplateField HeaderText="No." HeaderStyle-Width="40" ItemStyle-HorizontalAlign="Right" >
                <ItemTemplate>
                    <%# Container.DataItemIndex + 1 %>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField HeaderText="Programa" DataField="Programa" HeaderStyle-Width="300" />
            <asp:BoundField HeaderText="Género" DataField="Genero" HeaderStyle-Width="200" />
            <asp:BoundField HeaderText="M.Fílmico?" DataField="FlagMaterial" HeaderStyle-Width="100" />
            <asp:BoundField HeaderText="Grupo Programa" DataField="GrupoPrograma" HeaderStyle-Width="300" />
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

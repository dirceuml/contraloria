<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Pasivos.aspx.vb" Inherits="ATVContraloria.Pasivos" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
    <div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Consultas : </span><span class="fontTitulo1b">Pasivos</span></td>
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
    <td height="5"></td>
</tr>
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:440px;text-align:center;overflow:auto;">
    <asp:GridView ID="gdvPasivos" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" 
        GridLines="Both" DataKeyNames="IdGrupoPasivos" ShowFooter="True">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Seccion" ItemStyle-HorizontalAlign="Left" HeaderStyle-Width="200" >
                <ItemTemplate>
                    &nbsp;<%# Eval("Seccion")%>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboSeccion" runat="server" Width="180" Font-Size="Smaller" ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboSeccion" runat="server" Width="180" Font-Size="Smaller" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Grupo" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Left" >
                <ItemTemplate>
                    &nbsp;<%# Eval("Grupo")%>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboGrupo" runat="server" Width="180" Font-Size="Smaller" ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboGrupo" runat="server" Width="180" Font-Size="Smaller" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="SubGrupo" ItemStyle-HorizontalAlign="Left" HeaderStyle-Width="200" >
                <ItemTemplate>
                    &nbsp;<%# Eval("SubGrupo")%>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboSubGrupo" runat="server" Width="180" Font-Size="Smaller"  ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboSubGrupo" runat="server" Width="180" Font-Size="Smaller" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cuenta" ItemStyle-HorizontalAlign="Left" HeaderStyle-Width="300" >
                <ItemTemplate>
                    &nbsp;<%# Eval("CodCuenta")%>&nbsp;<%# Eval("Cuenta")%>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:DropDownList ID="cboCuenta" runat="server" Width="280" Font-Size="Smaller"  ></asp:DropDownList>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList ID="cboCuenta" runat="server" Width="280" Font-Size="Smaller" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Gpo.ATV?" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="60" >
                <ItemTemplate>
                    <%# Eval("FlagGpoATV")%>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:CheckBox ID="chkFlagGpoATV" runat="server" CssClass="InputText10" Checked='<%# If(Eval("FlagGpoATV").ToString() = "S", True, False) %>' />
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkFlagGpoATV" runat="server" CssClass="InputText10"  />                
                </FooterTemplate>
            </asp:TemplateField>            
            <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="60" >
                <ItemTemplate>
                    <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" CausesValidation="False" ><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>&nbsp;
                    <span onclick="return confirm('Esta usted seguro en borrar el registro?')"><asp:LinkButton ID="lnkDelete" runat="server" CommandName="Delete"><asp:Image ID="ImgDelete" runat="server" ImageUrl="images/delete.gif" AlternateText="Eliminar" /></asp:LinkButton></span>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:LinkButton ID="lnkUpdate" runat="server" CausesValidation="True" CommandName="Update"><asp:Image ID="imgUpdate" runat="server" ImageUrl="images/save.gif" AlternateText="Grabar" /></asp:LinkButton>&nbsp;
                        &nbsp;
                        <asp:LinkButton ID="lnkCancel" runat="server" CausesValidation="False" CommandName="Cancel"><asp:Image ID="imgCancel" runat="server" ImageUrl="images/cancel.gif" AlternateText="Cancelar" /></asp:LinkButton>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:LinkButton ID="lnkInsert" runat="server" OnClick="lnkNuevo_Click" CausesValidation="False" ><asp:Image ID="imgInsert" runat="server" ImageUrl="images/new.gif" AlternateText="Crear" /></asp:LinkButton>
                </FooterTemplate>
            </asp:TemplateField>
        </Columns>
        <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#DCDCDC" />
    </asp:GridView>
    <asp:GridView ID="gdvPasivosExp" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" EnableModelValidation="True" ForeColor="#333333" 
            GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
            <asp:BoundField HeaderText="Sección" DataField="Seccion" ItemStyle-Width="200px"/>
            <asp:BoundField HeaderText="Grupo" DataField="Grupo" ItemStyle-Width="200px"/>
            <asp:BoundField HeaderText="SubGrupo" DataField="SubGrupo" ItemStyle-Width="200px"/>
            <asp:BoundField HeaderText="Cod.Cuenta" DataField="CodCuenta" ItemStyle-Width="80px" />
            <asp:BoundField HeaderText="Cuenta" DataField="Cuenta" ItemStyle-Width="360px" />
            <asp:BoundField HeaderText="¿Gpo.ATV?" DataField="FlagGpoATV" ItemStyle-Width="60px" />
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


<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Ajustes.aspx.vb" Inherits="ATVContraloria.Ajustes" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
    <script type="text/javascript">
    <!--
    function Filtrar(objCombo) {
        var nId = objCombo.value;
        document.forms[0].ctl00$contenido$hdfAccion.value = "Recargar";
        document.forms[0].submit();
    }
    function ValidaConsulta() {
        document.forms[0].ctl00$contenido$hdfAccion.value = "Recargar";
        return true;
    }
    //-->
    </script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">
<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Actualización : </span><span class="fontTitulo1b">Ajustes</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr style="background-color:#4d81b3; height:22px;">
    <td align="left">
        <font class="fontTextBlanco"><b>Periodo:</b></font> <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="100px"></asp:DropDownList>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <font class="fontTextBlanco"><b>Tipo:</b></font> <asp:DropDownList ID="cboCuentaFC2" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="100px">
        <asp:ListItem Value="" Text="--TODO--"></asp:ListItem>
        <asp:ListItem Value="CONTABILIDAD" Text="CONTABILIDAD"></asp:ListItem>
        <asp:ListItem Value="TESORERIA" Text="TESORERIA"></asp:ListItem>
        </asp:DropDownList>
    </td>
    <td align="right"><asp:Button ID="btnDescargaExcel" runat="server" Text="Descargar a Excel" onclick="btnDescargaExcel_Click" Height="20px" Width="120px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" /></td>
</tr>
</table>
<asp:HiddenField ID="hdfAccion" Value="" runat="server" />
<br />
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvAjuste" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="IdAjuste" 
        onrowdatabound="gdvAjuste_RowDataBound"
        onrowediting="gdvAjuste_RowEditing" 
        OnRowDeleting="gdvAjuste_RowDeleting"
        onrowcancelingedit="gdvAjuste_RowCancelingEdit" 
        onrowupdating="gdvAjuste_RowUpdating">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" ItemStyle-CssClass="fontText10" ReadOnly ></asp:BoundField>
            <asp:TemplateField HeaderText="Fecha" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblFecha" runat="server" Text="lblFecha" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtFecha" runat="server" CssClass="InputText10" Width="85" MaxLength="10"></asp:TextBox></EditItemTemplate>
                <FooterTemplate><asp:TextBox ID="txtFechaNew" runat="server" CssClass="InputText10" Width="85" MaxLength="10"></asp:TextBox></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cuenta" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblCuentaFC2" runat="server" Text="lblCuentaFC2" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:DropDownList ID="cboCodCuentaFC2" runat="server" CssClass="fontText10" ></asp:DropDownList></EditItemTemplate>
                <FooterTemplate><asp:DropDownList ID="cboCodCuentaFC2New" runat="server" CssClass="fontText10" ></asp:DropDownList></FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="MontoUSD" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><%# Eval("MontoUSD", "{0:n2}")%></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoUSD" runat="server" Text='<%# Bind("MontoUSD")%>' CssClass="fontText10" ></asp:TextBox></EditItemTemplate>
                <FooterTemplate><asp:TextBox ID="txtMontoUSDNew" runat="server" Text='' CssClass="fontText10" ></asp:TextBox></FooterTemplate>
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
    <asp:GridView ID="gdvResultadosExp" runat="server" AutoGenerateColumns="False" Visible="false" >
        <Columns>
                <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" />
                <asp:BoundField HeaderText="Fecha" DataField="Fecha" DataFormatString="{0:yyyy-MM-dd}" />
                <asp:BoundField HeaderText="Codigo" DataField="Cuenta2"></asp:BoundField>
                <asp:BoundField HeaderText="MontoUSD" DataField="MontoUSD" />
        </Columns>
        <HeaderStyle BackColor="Silver" />
        <FooterStyle BackColor="Silver" Font-Bold="true" />
    </asp:GridView>
    </div> 
    </td>
</tr>
</table>
</asp:Content>
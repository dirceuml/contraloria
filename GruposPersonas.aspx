<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="GruposPersonas.aspx.vb" Inherits="ATVContraloria.GruposPersonas" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
    <script type="text/javascript">
    <!--
    function Filtrar(objCombo) {
        var nId = objCombo.value;
        document.forms[0].ctl00$contenido$hdfAccion.value = "Recargar";
        document.forms[0].submit();
    }
    function ValidaConsulta() {
        var sPersona = document.forms[0].ctl00$contenido$txtPersona.value;

        if (f_IsEmpty(sPersona)) {
            alert("Por favor, ingresar información en el campo persona");
            document.forms[0].ctl00$contenido$txtPersona.focus();
            return false;
        }
        else {
            document.forms[0].ctl00$contenido$hdfAccion.value = "Recargar";
            return true;
        }
    }
    //-->
    </script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Mantenimientos : </span><span class="fontTitulo1b">Agrupación Personas (Nuevo formato)</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>

</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
    <asp:HiddenField ID="hdfAccion" Value="" runat="server" />
    <asp:HiddenField ID="hdfNroRegistros" Value="0" runat="server" />
    <asp:HiddenField ID="hdfPagina" Value="" runat="server" />
    <br />
    <table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
    <tr><td style="background-color:#4d81b3;" colspan="5"><font class="fontTextBlanco"><b>Filtros</b></font></td></tr>
    <tr class="TDList1">
        <td><font class="fontTextNormal">Persona: </font></td>
        <td><asp:TextBox ID="txtPersona" runat="server" CssClass="InputText10" Width="100" MaxLength="10"></asp:TextBox></td>
        <td><font class="fontTextNormal">Grupos</font></td>
        <td><asp:DropDownList ID="cboGrupo" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="180px" ></asp:DropDownList></td>
        <td><asp:Button ID="btnConsultar" runat="server" Text="Consultar" OnClick="btnConsultar_Click" OnClientClick="javascript:if (!ValidaConsulta() ) return false;" Height="20px" CssClass ="fontBoton" /></td>
    </tr>
    <tr><td style="width:1px; height:2px;" colspan="5"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr><td style="background-color:#4d81b3;" colspan="5"><font class="txtBlanco"><b>
        <asp:Label ID="lblNumRegistros" runat="server" Text=""></asp:Label></b></font></td>
    </tr>
    <tr><td style="width:1px; height:2px;" colspan="5"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr>
        <td colspan="5">
        <asp:GridView ID="gdvPersona" runat="server" AutoGenerateColumns="False" Width="100%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" AllowPaging="True" DataKeyNames="CodPersona" 
            GridLines="Both" onpageindexchanging="gdvPersona_PageIndexChanging" 
            onrowdatabound="gdvPersona_RowDataBound" 
            onrowediting="gdvPersona_RowEditing" 
            onrowcancelingedit="gdvPersona_RowCancelingEdit" 
            onrowupdating="gdvPersona_RowUpdating" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Código" DataField="CodPersona" ReadOnly="true" ControlStyle-CssClass="fontText10" HeaderStyle-Width="50" ></asp:BoundField>
                <asp:BoundField HeaderText="Persona" DataField="Persona" ItemStyle-HorizontalAlign="left" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:TemplateField HeaderText="Grupo" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:Label ID="lnkPersonaN" runat="server" Text="lnkPersonaN" CssClass="fontText10" ></asp:Label>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:DropDownList ID="cboPersonaN" runat="server" CssClass="fontText10" ></asp:DropDownList>
                        <asp:TextBox ID="txtPersonaN" runat="server" CssClass="InputText10" Width="100" MaxLength="10"></asp:TextBox>
                    </EditItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="120" >
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
        </td>
    </tr>
    </table>

</asp:Content>
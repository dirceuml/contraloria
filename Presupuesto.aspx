<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Presupuesto.aspx.vb" Inherits="ATVContraloria.Presupuesto" %>
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
      <td width="61%"><span class="fontTitulo1">Módulo : </span><span class="fontTitulo1b">Presupuesto</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr style="background-color:#4d81b3; height:22px;">
    <td align="left"><font class="fontTextBlanco"><b>Año:</b></font> <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="100px"></asp:DropDownList>&nbsp;&nbsp;</td>
</tr>
</table>
<asp:HiddenField ID="hdfAccion" Value="" runat="server" />
<asp:HiddenField ID="hdfNroRegistros" Value="0" runat="server" />
<br />
<table border="0" cellspacing="0" cellpadding="5" width="90%" align="center">
<tr>
    <td>
    <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
    <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
    <asp:GridView ID="gdvFlujo" runat="server" AutoGenerateColumns="False" Width="100%"
        BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
        GridLines="Both" DataKeyNames="CodDetalle" 
        onrowdatabound="gdvFlujo_RowDataBound"
        onrowediting="gdvFlujo_RowEditing" 
        onrowcancelingedit="gdvFlujo_RowCancelingEdit" 
        onrowupdating="gdvFlujo_RowUpdating">                                
        <RowStyle BackColor="#FFFFFF" />
        <Columns>
            <asp:TemplateField HeaderText="Sección" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10">
                <ItemTemplate><%# Eval("CodSeccion")%></ItemTemplate>
                <EditItemTemplate><%# Eval("CodSeccion")%></EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Rubro" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10">
                <ItemTemplate><%# Eval("Rubro")%></ItemTemplate>
                <EditItemTemplate><%# Eval("Rubro")%><asp:HiddenField ID="hdfRubro" Value="0" runat="server" /></EditItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Enero" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoEnero" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoEnero" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Febrero" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoFebrero" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoFebrero" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Marzo" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoMarzo" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoMarzo" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Abril" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoAbril" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoAbril" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Mayo" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoMayo" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoMayo" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Junio" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoJunio" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoJunio" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Julio" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoJulio" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoJulio" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Agosto" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoAgosto" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoAgosto" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Setiembre" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoSetiembre" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoSetiembre" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Octubre" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoOctubre" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoOctubre" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Noviembre" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoNoviembre" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoNoviembre" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Diciembre" ItemStyle-HorizontalAlign="right" ControlStyle-CssClass="fontText10" >
                <ItemTemplate><asp:Label ID="lblMontoNetoDiciembre" runat="server" Text="lblMontoNeto" CssClass="fontText10" ></asp:Label></ItemTemplate>
                <EditItemTemplate><asp:TextBox ID="txtMontoNetoDiciembre" runat="server" Text='txtMontoNeto' CssClass="fontText10" Width="50" ></asp:TextBox></EditItemTemplate>   
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Acción" ItemStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" CausesValidation="False"><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                </ItemTemplate>
                <EditItemTemplate>
                    <asp:LinkButton ID="lnkUpdate" runat="server" CausesValidation="True" CommandName="Update"><asp:Image ID="imgUpdate" runat="server" ImageUrl="images/save.gif" AlternateText="Grabar" /></asp:LinkButton>&nbsp;
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
    <br />
    </div> 
    </td>
</tr>
</table>
</asp:Content>

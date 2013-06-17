<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ImportacionVentas.aspx.vb" Inherits="ATVContraloria.ImportacionVentas" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Importación SISMED : </span><span class="fontTitulo1b">Ventas</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
    
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">

<AjaxToolKit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" AsyncPostBackTimeout="3600">
</AjaxToolKit:ToolkitScriptManager>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
<ContentTemplate>
1) Seleccionar Periodo: <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList><br/><br/>
2) Importar Datos:  <asp:Button ID="btnImportaVentas" runat="server" Text="Importar Datos Ventas" Height="20px" Width="160px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" /><br/><br/>
3) Descargar Formato Rating: <asp:LinkButton ID="lnkDescargaRating" runat="server" Font-Size="Small">Descargar Formato</asp:LinkButton><br/><br/>
4) Importar Rating: <asp:FileUpload ID="fupRating" runat="server" />&nbsp;<asp:LinkButton ID="lnkImportaRating" runat="server" Font-Size="Small">Importar</asp:LinkButton><br/><br/>
<asp:Label ID="lblEstado" runat="server" Text="" ForeColor="Red"></asp:Label>
</ContentTemplate>
<Triggers>
    <asp:PostBackTrigger ControlID="lnkDescargaRating" />
    <asp:PostBackTrigger ControlID="lnkImportaRating" />
</Triggers>
</asp:UpdatePanel>

<asp:UpdateProgress DynamicLayout="false" ID="UpdateProgress1" runat="server">
    <ProgressTemplate>
        <img src="images/ajax-loader.gif" alt="" />
    </ProgressTemplate>
 </asp:UpdateProgress>
<AjaxToolKit:AlwaysVisibleControlExtender ID="AlwaysVisibleControlExtender1" runat="server" TargetControlID="UpdateProgress1" HorizontalSide="Center" VerticalSide="Middle" >
</AjaxToolKit:AlwaysVisibleControlExtender>

</asp:Content>


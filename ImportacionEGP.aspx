<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ImportacionEGP.aspx.vb" Inherits="ATVContraloria.ImportacionEGP" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Importación SISMED : </span><span class="fontTitulo1b">EGP</span></td>
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

<asp:Button ID="btnImportaEGP" runat="server" Text="Importar Datos EGP" Height="20px" Width="160px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" /><br/><br/>

<asp:Label ID="lblEstado" runat="server" Text="" ForeColor="Red"></asp:Label>
</ContentTemplate>
</asp:UpdatePanel>

<asp:UpdateProgress DynamicLayout="false" ID="UpdateProgress1" runat="server">
    <ProgressTemplate>
        <img src="images/ajax-loader.gif" alt="" />
    </ProgressTemplate>
 </asp:UpdateProgress>
<AjaxToolKit:AlwaysVisibleControlExtender ID="AlwaysVisibleControlExtender1" runat="server" TargetControlID="UpdateProgress1" HorizontalSide="Center" VerticalSide="Middle" >
</AjaxToolKit:AlwaysVisibleControlExtender>

</asp:Content>
<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ImportacionEvolucionVentas.aspx.vb" Inherits="ATVContraloria.ImportacionEvolucionVentas" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Importación SISMED : </span><span class="fontTitulo1b">Evolución Ventas</span></td>
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
Importar Evolución Ventas. Periodo: <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10"></asp:DropDownList>&nbsp;
<asp:Button ID="btnImportaEvolucionVentas" runat="server" Text="Importar Datos" Height="20px" Width="100px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" /><br/>
<asp:Label ID="lblEstado" runat="server" Text="" ForeColor="Red"></asp:Label><br/>
Descargar Formato <b>Presupuesto Facturación</b>. Año: <asp:DropDownList ID="cboAño" runat="server" CssClass="InputText10"></asp:DropDownList>&nbsp;
<asp:LinkButton ID="lnkDescargaVentaPpto" runat="server" Font-Size="Small">Descargar Formato</asp:LinkButton><br/><br/>
Importar <b>Presupuesto Facturación</b>. Seleccione archivo: <asp:FileUpload ID="fupVentaPpto" runat="server" />&nbsp;
<asp:LinkButton ID="lnkImportaVentaPpto" runat="server" Font-Size="Small">Importar Presupuesto</asp:LinkButton><br/>
<asp:Label ID="lblEstadoPpto" runat="server" Text="" ForeColor="Red"></asp:Label><br/>
Descargar Archivo <b>Ventas Global y ATV_SUR</b>.&nbsp;
<asp:LinkButton ID="lnkDescargaVenta2" runat="server" Font-Size="Small">Descargar Archivo</asp:LinkButton><br/><br/>
Importar Archivo <b>Ventas Global y ATV_SUR</b>. Seleccione archivo: <asp:FileUpload ID="fupVenta2" runat="server" />&nbsp;
<asp:LinkButton ID="lnkImportaVenta2" runat="server" Font-Size="Small">Importar Ventas</asp:LinkButton><br/>
<asp:Label ID="lblEstado2" runat="server" Text="" ForeColor="Red"></asp:Label><br/>
</ContentTemplate>
<Triggers>
    <asp:PostBackTrigger ControlID="lnkDescargaVentaPpto" />
    <asp:PostBackTrigger ControlID="lnkImportaVentaPpto" />
    <asp:PostBackTrigger ControlID="lnkDescargaVenta2" />
    <asp:PostBackTrigger ControlID="lnkImportaVenta2" />
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
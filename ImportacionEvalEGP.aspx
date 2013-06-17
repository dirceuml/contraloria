<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ImportacionEvalEGP.aspx.vb" Inherits="ATVContraloria.ImportacionEvalEGP" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Importación : </span><span class="fontTitulo1b">Evaluación EGP</span></td>
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
Descargar Formato <b>Evaluacion EGP Datos</b>. Año: <asp:DropDownList ID="cboAño" runat="server" CssClass="InputText10"></asp:DropDownList>&nbsp;
<asp:LinkButton ID="lnkDescargaEvalEGP" runat="server" Font-Size="Small">Descargar Formato</asp:LinkButton><br/><br/>
Importar <b>Evaluación EGP Datos</b>. Seleccione archivo: <asp:FileUpload ID="fupEvalEGP" runat="server" />&nbsp;
<asp:LinkButton ID="lnkImportaEvalEGP" runat="server" Font-Size="Small">Importar Evaluación</asp:LinkButton><br/>
<asp:Label ID="lblEstado" runat="server" Text="" ForeColor="Red"></asp:Label><br/>
</ContentTemplate>
<Triggers>
    <asp:PostBackTrigger ControlID="lnkDescargaEvalEGP" />
    <asp:PostBackTrigger ControlID="lnkImportaEvalEGP" />
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
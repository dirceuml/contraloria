<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="AjustesExt.aspx.vb" Inherits="ATVContraloria.AjustesExt" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxToolKit" %>
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
      <td width="61%"><span class="fontTitulo1">Actualización : </span><span class="fontTitulo1b">Ajustes (Exterior)</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
<AjaxToolKit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></AjaxToolKit:ToolkitScriptManager>
<asp:UpdatePanel ID="UpdatePanel2" runat="server">
<ContentTemplate>
<table border="0" cellspacing="0" cellpadding="5" width="100%" align="center">
<tr style="background-color:#4d81b3; height:22px;">
    <td align="left"><font class="fontTextBlanco"><b>Periodo:</b></font> <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="100px"></asp:DropDownList>&nbsp;&nbsp;</td>
    <td align="right">
        <asp:Button ID="btnNuevo" runat="server" Text="Nuevo" OnClick="btnNuevo_Click" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" Width="100px" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnDescargaExcel" runat="server" Text="Descargar a Excel" onclick="btnDescargaExcel_Click" Width="120px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
    </td>
</tr>
</table>
</ContentTemplate>
<Triggers>
    <asp:PostBackTrigger ControlID="btnDescargaExcel" />
</Triggers>
</asp:UpdatePanel>
<br />
<table border="0" cellspacing="0" cellpadding="5" width="100%" align="center">
<tr>
    <td>
        <asp:UpdatePanel ID="upnlResumen" runat="server">
        <ContentTemplate>
        <asp:HiddenField ID="hdfAccion" Value="" runat="server" />
        <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
        <asp:Label ID="lblMensaje" runat="server" Text="" CssClass="fontTextRojo10"></asp:Label>
        <asp:GridView ID="gdvAjusteExt" runat="server" AutoGenerateColumns="False" Width="100%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
            GridLines="Both" DataKeyNames="IdAjuste" 
            OnRowDeleting="gdvAjusteExt_RowDeleting"
            onrowdatabound="gdvAjusteExt_RowDataBound">                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" ItemStyle-CssClass="fontText10" Visible="false" ></asp:BoundField>
                <asp:BoundField HeaderText="Fecha" DataField="Fecha" ItemStyle-HorizontalAlign="center" ItemStyle-CssClass="fontText10" DataFormatString="{0:yyyy-MM-dd}" ItemStyle-Width="80" />
                <asp:TemplateField HeaderText="Tipo Ajuste" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate><asp:Label ID="lblCodTipoAjuste" runat="server" Text="lblCodTipoAjuste" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Cuenta Origen" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblCuentaFCOrig" runat="server" Text="lblCuentaFCOrig" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Persona Origen" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblPersonaOrig" runat="server" Text="lblPersonaOrig" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Cuenta Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblCuentaFCDes" runat="server" Text="lblCuentaFCDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Persona Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblPersonaDes" runat="server" Text="lblPersonaDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Area Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblAreaDes" runat="server" Text="lblAreaDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="MontoUSD" DataField="MontoUSD" DataFormatString="{0:n2}" ItemStyle-HorizontalAlign="Right" ItemStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Observación" DataField="Observacion" ItemStyle-CssClass="fontText10" ItemStyle-HorizontalAlign="left" ></asp:BoundField>
                <asp:TemplateField HeaderText="Editar" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="50">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Editar" CommandArgument='<%# Eval("IdAjuste") %>'><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                        &nbsp;
                        <span onclick="return confirm('Esta usted seguro en borrar el registro?')"><asp:LinkButton ID="lnkDelete" runat="server" CommandName="Delete"><asp:Image ID="ImgDelete" runat="server" ImageUrl="images/delete.gif" AlternateText="Eliminar" /></asp:LinkButton></span>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        <br />
        <asp:GridView ID="gdvResultadosExp" runat="server" AutoGenerateColumns="False" Visible="false" onrowdatabound="gdvResultadosExp_RowDataBound">
        <Columns>
                <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" ItemStyle-CssClass="fontText10" />
                <asp:BoundField HeaderText="Fecha" DataField="Fecha" ItemStyle-HorizontalAlign="center" ItemStyle-CssClass="fontText10" DataFormatString="{0:yyyy-MM-dd}" />
                <asp:TemplateField HeaderText="Tipo Ajuste" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblCodTipoAjuste" runat="server" Text="lblCodTipoAjuste" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Cuenta Origen" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblCuentaFCOrig" runat="server" Text="lblCuentaFCOrig" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Persona Origen" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblPersonaOrig" runat="server" Text="lblPersonaOrig" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Cuenta Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblCuentaFCDes" runat="server" Text="lblCuentaFCDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Persona Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblPersonaDes" runat="server" Text="lblPersonaDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Area Destino" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lblAreaDes" runat="server" Text="lblAreaDes" CssClass="fontText10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="Observación" DataField="Observacion"></asp:BoundField>
                <asp:BoundField HeaderText="MontoUSD" DataField="MontoUSD" />
        </Columns>
        <HeaderStyle BackColor="Silver" />
        <FooterStyle BackColor="Silver" Font-Bold="true" />
    </asp:GridView>
        </div> 
        </ContentTemplate>
        <Triggers><asp:AsyncPostBackTrigger ControlID="btnSaveDetalle" /></Triggers>
        </asp:UpdatePanel>
    </td>
</tr>
</table>

<%--Detalle------------------------------------------------------------------------------------------%>
    <asp:Button ID="btnInvisible1" runat="server" Text="Invisible" style="display:none;" />
    <asp:Panel ID="pnlDetalle" runat="server" Width="800px" BorderWidth="0px" CssClass="modalPopup" BackColor="#FFFFFF">
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
            <asp:HiddenField ID="hdfIdAjuste" Value="" runat="server" />
            <table cellpadding="0" cellspacing="1" border="0" width="100%">
            <tr>
                <td rowspan="20"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
                <td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
                <td rowspan="20"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
            </tr>
            <tr style="background-color:#4d81b3;"><td colspan="4" align="left"><font class="fontTextBlanco"><b>Datos de la Cuenta</b></font></td></tr>
            <%--<tr class="TDList0">
                <td><font class="fontText10">Periodo</font></td>
                <td colspan="3"><asp:DropDownList ID="cboIdPeriodo" runat="server" CssClass="fontTextSmall" ></asp:DropDownList></td>
            </tr>--%>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Fecha</font></td>
                <td><asp:TextBox ID="txtFecha" runat="server" CssClass="InputText10" Width="85" MaxLength="10"></asp:TextBox>&nbsp;&nbsp;<font class="fontText10">[aaaa-mm-dd]</font></td>
                <td><font class="fontText10">Tipo Ajuste</font></td>
                <td>
                    <asp:RadioButton ID="rdbIngreso" GroupName="rdbCodTipoAjuste" Text="Adicion" runat="server" Checked="true" CssClass="combos" />&nbsp;
                    <asp:RadioButton ID="rdbEgreso" GroupName="rdbCodTipoAjuste" Text="Cambio" runat="server" CssClass="combos" />
                </td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Cuenta Origen</font></td>
                <td colspan="3"><asp:DropDownList ID="cboCodCuentaFCOrig" runat="server" CssClass="InputText10" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Persona Origen</font></td>
                <td colspan="3"><asp:DropDownList ID="cboCodPersonaOrig" runat="server" CssClass="InputText10" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Cuenta Destino</font></td>
                <td colspan="3"><asp:DropDownList ID="cboCodCuentaFCDes" runat="server" CssClass="InputText10" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Persona Destino</font></td>
                <td colspan="3"><asp:DropDownList ID="cboCodPersonaDes" runat="server" CssClass="InputText10" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Area Destino</font></td>
                <td colspan="3"><asp:DropDownList ID="cboCodAreaDes" runat="server" CssClass="InputText10" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Observación</font></td>
                <td colspan="3"><asp:TextBox ID="txtObservacion" runat="server" Text='' CssClass="InputText10" MaxLength="80" Width="400" ></asp:TextBox></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">MontoUSD</font></td>
                <td colspan="3"><asp:TextBox ID="txtMontoUSD" runat="server" Text='' CssClass="InputText10" ></asp:TextBox></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr>
                <td colspan="4" align="center">
                    <asp:Button ID="btnSaveDetalle" runat="server" Text="Grabar" OnClick="btnSaveDetalle_Click" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" Width="100px" />&nbsp;&nbsp;
                    <asp:Button ID="btnCloseDetalle" runat="server" Text="Salir" OnClick="btnCloseDetalle_Click" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" Width="100px" />
                </td>
            </tr>
            <tr><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            </table>
            </ContentTemplate>
            <Triggers><asp:AsyncPostBackTrigger ControlID="btnCloseDetalle" /></Triggers>
        </asp:UpdatePanel>  
    </asp:Panel>
    <AjaxToolKit:ModalPopupExtender ID="mpupDetalle" runat="server" TargetControlID="btnInvisible1" PopupControlID="pnlDetalle" BackgroundCssClass="modalBackground" PopupDragHandleControlID="pnlDetalle"></AjaxToolKit:ModalPopupExtender>

    <asp:UpdateProgress DynamicLayout="false" ID="UpdateProgress1" runat="server">
        <ProgressTemplate>
            <img src="images/ajax-loader.gif" alt="" />
        </ProgressTemplate>
     </asp:UpdateProgress>
    <AjaxToolKit:AlwaysVisibleControlExtender ID="AlwaysVisibleControlExtender1" runat="server" TargetControlID="UpdateProgress1" HorizontalSide="Center" VerticalSide="Middle" >
    </AjaxToolKit:AlwaysVisibleControlExtender>

</asp:Content>

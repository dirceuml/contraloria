<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="Actualizacion.aspx.vb" Inherits="ATVContraloria.Actualizacion" EnableEventValidation ="false" %>
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
      <td width="61%"><span class="fontTitulo1">Actualización : </span><span class="fontTitulo1b">Cuentas Flujo Caja</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>

</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">
    <AjaxToolKit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></AjaxToolKit:ToolkitScriptManager>
    <asp:HiddenField ID="hdfAccion" Value="" runat="server" />
    <asp:HiddenField ID="hdfNroRegistros" Value="0" runat="server" />
    <br />
    <table border="0" cellspacing="0" cellpadding="5" width="95%" align="center">
    <tr style="background-color:#4d81b3; height:22px;">
        <td colspan="3" align="left"><font class="fontTextBlanco"><b>Filtros</b></font></td>
        <td colspan="3" align="right"><font class="fontTextBlanco"><b>Periodo:</b></font> <asp:DropDownList ID="cboPeriodo" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="100px"></asp:DropDownList>&nbsp;&nbsp;</td>
    </tr>
    <tr class="TDList1">
        <td><font class="fontTextSmall">Cta. Banco</font></td>
        <td><asp:DropDownList ID="cboCuentaBanco" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px"></asp:DropDownList></td>
        <td><font class="fontTextSmall">Cta. Caja</font></td>
        <td><asp:DropDownList ID="cboCuentaFC" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px" ></asp:DropDownList></td>
        <td><font class="fontTextSmall">Cta. Nueva</font></td>
        <td><asp:DropDownList ID="cboCuentaFCNueva" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px" ></asp:DropDownList></td>
    </tr>
    <tr class="TDList0">
        <td><font class="fontTextSmall">Area</font></td>
        <td><asp:DropDownList ID="cboArea" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px"></asp:DropDownList></td>
        <td><font class="fontTextSmall">Area Nueva</font></td>
        <td><asp:DropDownList ID="cboAreaNueva" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px"></asp:DropDownList></td>
        <td><font class="fontTextSmall">Persona</font></td>
        <td><asp:DropDownList ID="cboPersona" runat="server" CssClass="InputText10" onchange="Filtrar(this);" Width="280px"></asp:DropDownList></td>
    </tr>
    <tr class="TDList1">
        <td><font class="fontTextSmall">Persona</font></td>
        <td>
            <asp:TextBox ID="txtPersona" runat="server" CssClass="InputText10" Width="120" MaxLength="10"></asp:TextBox>&nbsp;&nbsp;&nbsp;
            <font class="fontTextSmall">Glosa</font>&nbsp;
            <asp:TextBox ID="txtGlosa" runat="server" CssClass="InputText10" Width="106" MaxLength="10"></asp:TextBox>
        </td>
        <td><font class="fontTextSmall">Fecha Voucher</font></td>
        <td>
            <asp:TextBox ID="txtFecha" runat="server" CssClass="InputText10" Width="85" MaxLength="10"></asp:TextBox>&nbsp;<img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFecha,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" />&nbsp;&nbsp;&nbsp;
            <font class="fontTextSmall">Nro Voucher</font>&nbsp;
            <asp:TextBox ID="txtNroVoucher" runat="server" CssClass="InputText10" Width="85" MaxLength="10"></asp:TextBox>
        </td>
        <td></td>
        <td>
            <asp:Button ID="btnConsultar" runat="server" Text="Consultar" OnClick="btnConsultar_Click" OnClientClick="javascript:if (!ValidaConsulta() ) return false;" Height="20px" Width="120px" CssClass ="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
            <asp:Button ID="btnDescargaExcel" runat="server" Text="Descargar a Excel" onclick="btnDescargaExcel_Click" Height="20px" Width="120px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
        </td>
    </tr>
    </table>

    <table border="0" cellspacing="0" cellpadding="0" width="95%" align="center">
    <tr><td style="width:1px; height:2px;"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr><td style="background-color:#4d81b3;"><font class="txtBlanco"><b><asp:Label ID="lblNumRegistros" runat="server" Text=""></asp:Label></b></font></td></tr>
    <tr><td style="width:1px; height:2px;"><img src="images/trans.gif" style="width:1px; height:2px;" alt="" /></td></tr>
    <tr>
        <td>
        <asp:UpdatePanel ID="upnlResumen" runat="server">
        <ContentTemplate>
        <asp:HiddenField ID="hdfPagina" Value="" runat="server" />
        <div style="background-color:#FFFFFF;height:400px;text-align:center;overflow:auto;">
        <asp:GridView ID="gdvMovimiento" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" Width="100%" GridLines="Both" DataKeyNames="IdMovimiento" 
            onpageindexchanging="gdvMovimiento_PageIndexChanging"
            onrowdatabound="gdvMovimiento_RowDataBound" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="Nº" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Bco." DataField="CodCuentaBanco" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Cta.Banco" DataField="CuentaBanco" Visible="false"></asp:BoundField>
                <asp:BoundField HeaderText="Vou" DataField="NroVoucher" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Item" DataField="NroItem" ItemStyle-HorizontalAlign="center" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Fecha" DataField="Fecha" Visible="false" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" DataFormatString="{0:dd/MM/yy}" ></asp:BoundField>
                <asp:BoundField HeaderText="CodCuentaFC" DataField="CodCuentaFC" ItemStyle-HorizontalAlign="left" Visible="false"></asp:BoundField>
                <asp:BoundField HeaderText="Cta.Caja" DataField="CuentaFC" ItemStyle-HorizontalAlign="left" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:TemplateField HeaderText="Cta.Nueva" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lnkCuentaDestino" runat="server" Text="lnkCuentaDestino" CssClass="fontTextRojo10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="CodPersona" DataField="CodPersona" ItemStyle-HorizontalAlign="left" Visible="false"></asp:BoundField>
                <asp:BoundField HeaderText="Persona" DataField="Persona" ItemStyle-HorizontalAlign="left" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Glosa" DataField="Glosa" ItemStyle-HorizontalAlign="left" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Observaciones" DataField="Observaciones" ItemStyle-HorizontalAlign="left" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="CodArea" DataField="CodArea" ItemStyle-HorizontalAlign="left" Visible="false"></asp:BoundField>
                <asp:BoundField HeaderText="Area" DataField="Area" ItemStyle-HorizontalAlign="left" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:TemplateField HeaderText="Area Nueva" ItemStyle-HorizontalAlign="left">
                    <ItemTemplate><asp:Label ID="lnkAreaNueva" runat="server" Text="lnkAreaNueva" CssClass="fontTextRojo10" ></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="Neto US$" DataField="MontoBaseUSD" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="IGV US$" DataField="MontoIGVUSD" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:BoundField HeaderText="Bruto US$" DataField="MontoUSD" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" ReadOnly="true" ControlStyle-CssClass="fontText10" ></asp:BoundField>
                <asp:TemplateField HeaderText="Asignar" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Editar" CommandArgument='<%# Eval("IdMovimiento") %>'><asp:Image ID="imgEdit" runat="server" ImageUrl="images/edit.gif" AlternateText="Editar" /></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="#333333" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        <asp:GridView ID="gdvResultadosExp" runat="server" AutoGenerateColumns="False" 
            Visible="false" onrowdatabound="gdvResultadosExp_RowDataBound">
            <Columns>
                    <asp:BoundField HeaderText="Periodo" DataField="IdPeriodo" />
                    <asp:BoundField HeaderText="Cod.Cta.Banco" DataField="CodCuentaBanco"></asp:BoundField>
                    <asp:BoundField HeaderText="Cta.Banco" DataField="CuentaBanco" />
                    <asp:BoundField HeaderText="Voucher" DataField="NroVoucher" />
                    <asp:BoundField HeaderText="Item" DataField="NroItem" />
                    <asp:BoundField HeaderText="Fecha" DataField="Fecha" DataFormatString="{0:yyyy-MM-dd}" />
                    <asp:BoundField HeaderText="Cod.Cta.Caja" DataField="CodCuentaFC"></asp:BoundField>
                    <asp:BoundField HeaderText="Cta.Caja" DataField="CuentaFC" />
                    <asp:BoundField HeaderText="Cta.Caja Nueva" DataField="CuentaFC2" />
                    <asp:BoundField HeaderText="CodPersona" DataField="CodPersona"></asp:BoundField>
                    <asp:BoundField HeaderText="Persona" DataField="Persona" />
                    <asp:BoundField HeaderText="Glosa" DataField="Glosa" />
                    <asp:BoundField HeaderText="Observaciones" DataField="Observaciones" />
                    <asp:BoundField HeaderText="Cod.Area" DataField="CodArea" />
                    <asp:BoundField HeaderText="Area" DataField="Area" />
                    <asp:BoundField HeaderText="Area Nueva" DataField="Area2" />
                    <asp:BoundField HeaderText="Cod.Movimiento" DataField="CodGrupoMov" />
                    <asp:BoundField HeaderText="Movimiento" DataField="GrupoMov" />
                    <asp:BoundField HeaderText="Neto US$" DataField="MontoBaseUSD" />
                    <asp:BoundField HeaderText="IGV US$" DataField="MontoIGVUSD" />
                    <asp:BoundField HeaderText="Bruto US$" DataField="MontoUSD" />
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
            <asp:HiddenField ID="hdfIdMovimiento" Value="" runat="server" />
            <table cellpadding="0" cellspacing="1" border="0" width="100%">
            <tr>
                <td rowspan="20"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
                <td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
                <td rowspan="20"><img src="images/trans.gif" width="10px" height="10px" alt=""></td>
            </tr>
            <tr style="background-color:#4d81b3;"><td colspan="4" align="left"><font class="fontTextBlanco"><b>Datos de la Cuenta</b></font></td></tr>
            <tr class="TDList1">
                <td><font class="fontText10">Periodo</font></td>
                <td><asp:Label ID="lblIdPeriodo" runat="server" CssClass="fontText10" Text="lblIdPeriodo"></asp:Label></td>
                <td><font class="fontText10">Fecha</font></td>
                <td><asp:Label ID="lblFecha" runat="server" CssClass="fontText10" Text="lblFecha" ></asp:Label></td>
            </tr>
            <tr class="TDList0">
                <td><font class="fontText10">Voucher</font></td>
                <td><asp:Label ID="lblNroVoucher" runat="server" CssClass="fontText10" Text="lblNroVoucher"></asp:Label></td>
                <td><font class="fontText10">Item</font></td>
                <td><asp:Label ID="lblNroItem" runat="server" CssClass="fontText10" Text="lblNroItem"></asp:Label></td>
            </tr>
            <tr class="TDList1">
                <td><font class="fontText10">Codigo Banco</font></td>
                <td><asp:Label ID="lblCodCuentaBanco" runat="server" CssClass="fontText10" Text="lblCodCuentaBanco"></asp:Label></td>
                <td><font class="fontText10">Cuenta.Banco</font></td>
                <td><asp:Label ID="lblCuentaBanco" runat="server" CssClass="fontText10" Text="lblCuentaBanco"></asp:Label></td>
            </tr>
            <tr class="TDList0">
                <td><font class="fontText10">Codigo Caja</font></td>
                <td><asp:Label ID="lblCodCuentaFC" runat="server" CssClass="fontText10" Text="lblCodCuentaFC"></asp:Label></td>
                <td><font class="fontText10">Cuenta.Caja</font></td>
                <td><asp:Label ID="lblCuentaFC" runat="server" CssClass="fontText10" Text="lblCuentaFC"></asp:Label></td>
            </tr>
            <tr class="TDList1">
                <td><font class="fontText10">Codigo Persona</font></td>
                <td><asp:Label ID="lblCodPersona" runat="server" CssClass="fontText10" Text="lblCodPersona"></asp:Label></td>
                <td><font class="fontText10">Persona</font></td>
                <td><asp:Label ID="lblPersona" runat="server" CssClass="fontText10" Text="lblPersona"></asp:Label></td>
            </tr>
            <tr class="TDList0">
                <td><font class="fontText10">Glosa</font></td>
                <td><asp:Label ID="lblGlosa" runat="server" CssClass="fontText10" Text="lblGlosa"></asp:Label></td>
                <td><font class="fontText10">Observaciones</font></td>
                <td><asp:Label ID="lblObservaciones" runat="server" CssClass="fontText10" Text="lblObservaciones"></asp:Label></td>
            </tr>
            <tr class="TDList1">
                <td><font class="fontText10">Codigo Area</font></td>
                <td><asp:Label ID="lblCodArea" runat="server" CssClass="fontText10" Text="lblCodArea"></asp:Label></td>
                <td><font class="fontText10">Area</font></td>
                <td><asp:Label ID="lblArea" runat="server" CssClass="fontText10" Text="lblArea"></asp:Label></td>
            </tr>
            <tr class="TDList0">
                <td><font class="fontText10">Codigo Movimiento</font></td>
                <td><asp:Label ID="lblCodGrupoMov" runat="server" CssClass="fontText10" Text="lblCodGrupoMov"></asp:Label></td>
                <td><font class="fontText10">Movimiento</font></td>
                <td><asp:Label ID="lblGrupoMov" runat="server" CssClass="fontText10" Text="lblGrupoMov"></asp:Label></td>
            </tr>
            <tr class="TDList1">
                <td><font class="fontText10">Neto US$</font></td>
                <td><asp:Label ID="lblMontoBaseUSD" runat="server" CssClass="fontText10" Text="lblMontoBaseUSD"></asp:Label></td>
                <td><font class="fontText10">IGV US$</font></td>
                <td><asp:Label ID="lblMontoIGVUSD" runat="server" CssClass="fontText10" Text="lblMontoIGVUSD"></asp:Label></td>
            </tr>
            <tr class="TDList0">
                <td><font class="fontText10">Bruto US$</font></td>
                <td><asp:Label ID="lblMontoUSD" runat="server" CssClass="fontText10" Text="lblMontoUSD"></asp:Label></td>
                <td><font class="fontText10">&nbsp;</font></td>
                <td><asp:Label ID="lblVacio" runat="server" CssClass="fontText10" Text="&nbsp;"></asp:Label></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr class="TDList0">
                <td><font class="fontText10">Cta.Nueva</font></td>
                <td><asp:DropDownList ID="cboCuentaFC2" runat="server" CssClass="fontTextSmall" ></asp:DropDownList></td>
                <td><font class="fontText10">Area Nueva</font></td>
                <td><asp:DropDownList ID="cboArea2" runat="server" CssClass="fontTextSmall" ></asp:DropDownList></td>
            </tr>
            <tr class="TDList1"><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            <tr>
                <td colspan="4" align="center">
                    <asp:Button ID="btnCloseDetalle" runat="server" Text="Salir" OnClick="btnCloseDetalle_Click" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" Width="100px" />
                    &nbsp;&nbsp;
                    <asp:Button ID="btnSaveDetalle" runat="server" Text="Grabar" OnClick="btnSaveDetalle_Click" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" Width="100px" />
                </td>
            </tr>
            <tr><td colspan="4"><img src="images/trans.gif" width="10px" height="10px" alt=""></td></tr>
            </table>
            </ContentTemplate>
            <Triggers><asp:AsyncPostBackTrigger ControlID="btnCloseDetalle" /></Triggers>
        </asp:UpdatePanel>  
    </asp:Panel>
    <AjaxToolKit:ModalPopupExtender ID="mpupDetalle" runat="server" TargetControlID="btnInvisible1" PopupControlID="pnlDetalle" BackgroundCssClass="modalBackground" PopupDragHandleControlID="pnlDetalle"></AjaxToolKit:ModalPopupExtender>
</asp:Content>

<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/ATVContraloria.Master" CodeBehind="ValidacionFCS.aspx.vb" Inherits="ATVContraloria.ValidacionFCS" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cabecera" runat="server">
    <script type="text/javascript">
    <!--
    function Filtrar(objCombo) {
        document.forms[0].ctl00$contenido$hdfAccion.value = "Recargar";
        document.forms[0].submit();
    }
    //-->
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="titulo" runat="server">

<div class="barraherram">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablaCabTitFech">
    <tr>
      <td width="61%"><span class="fontTitulo1">Importación : </span><span class="fontTitulo1b">Validación Flujo Caja Semanal</span></td>
      <td width="39%" align="right" nowrap="nowrap">&nbsp;</td>
    </tr>
  </table>
</div>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="contenido" runat="server">

    <asp:HiddenField ID="hdfAccion" Value="" runat="server" />
    <table border="0" cellspacing="0" cellpadding="5" width="100%" align="center">
    <tr><td colspan="4" style="background-color:#4d81b3;"><font class="fontTextBlanco"><b>&nbsp;&nbsp;Filtros</b></font></td></tr>
    <tr class="TDList1">
        <td colspan="4">&nbsp;
            <%--<font class="fontTextSmall">Periodo:</font>--%>
            <font class="fontTextSmall">Desde:</font><asp:TextBox ID="txtFechaIni" runat="server" CssClass="fontTextSmall" Width="60" MaxLength="10"></asp:TextBox><img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFechaIni,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" />&nbsp;
            <font class="fontTextSmall">Hasta:</font><asp:TextBox ID="txtFechaFin" runat="server" CssClass="fontTextSmall" Width="60" MaxLength="10"></asp:TextBox><img src="images/calendario.gif" onclick="displayCalendar(document.forms[0].ctl00$contenido$txtFechaFin,'yyyy-mm-dd',this)" border="0" width="23" height="19" alt="Calendario" />&nbsp;
            <font class="fontTextSmall">Cta.Banco:</font>
            <asp:DropDownList ID="cboCtaBanco" runat="server" CssClass="fontTextSmall" AutoPostBack="true"></asp:DropDownList>&nbsp;
            <asp:Button ID="btnConsultar" runat="server" Text="Consultar" OnClick="btnConsultar_Click" Height="20px" Width="80px" CssClass ="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
            <asp:Button ID="btnDescargaExcel" runat="server" Text="Excel" onclick="btnDescargaExcel_Click" Height="20px" Width="80px" CssClass="fontBoton" onmouseover="this.className='fontBotonp'" onmouseout="this.className='fontBoton'" />
        </td>
    </tr>
    <tr>
        <td colspan="4">
        <div style="background-color:#FFFFFF;height:480px;text-align:center;overflow:auto;">
        <asp:GridView ID="gdvCuenta" runat="server" AutoGenerateColumns="False" Width="100%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
            GridLines="Both" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="IdPeriodo" DataField="IdPeriodo" HeaderStyle-Width="80" Visible="false" ></asp:BoundField>
                <asp:BoundField HeaderText="Código" DataField="CodCuenta" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ></asp:BoundField>
                <asp:BoundField HeaderText="Cuenta" DataField="Cuenta" ItemStyle-HorizontalAlign="left"></asp:BoundField>                
                <asp:BoundField HeaderText="Monto Neto" DataField="MontoNeto" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="90" ></asp:BoundField>
                <asp:BoundField HeaderText="IGV" DataField="IGV" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="70" ></asp:BoundField>
                <asp:BoundField HeaderText="Monto Bruto" DataField="MontoBruto" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="90" ></asp:BoundField>
                <asp:BoundField HeaderText="M. Bruto (Con Signo)" DataField="MontoBruto2" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="120" ></asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        <asp:GridView ID="gdvCuenta2" runat="server" AutoGenerateColumns="False" Width="100%"
            BorderWidth="1px" BorderColor="#4d81b3" CellPadding="3" ShowFooter="true"
            GridLines="Both" >                                
            <RowStyle BackColor="#FFFFFF" />
            <Columns>
                <asp:BoundField HeaderText="IdPeriodo" DataField="IdPeriodo" HeaderStyle-Width="80" Visible="false" ></asp:BoundField>
                <asp:BoundField HeaderText="Código" DataField="CodCuenta" ItemStyle-HorizontalAlign="center" HeaderStyle-Width="50" ></asp:BoundField>
                <asp:BoundField HeaderText="Cuentas No Consideradas" DataField="Cuenta" ItemStyle-HorizontalAlign="left"></asp:BoundField>                
                <asp:BoundField HeaderText="Monto Neto" DataField="MontoNeto" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="90" ></asp:BoundField>
                <asp:BoundField HeaderText="IGV" DataField="IGV" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="70" ></asp:BoundField>
                <asp:BoundField HeaderText="Monto Bruto" DataField="MontoBruto" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="90" ></asp:BoundField>
                <asp:BoundField HeaderText="M. Bruto (Con Signo)" DataField="MontoBruto2" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:c}" HeaderStyle-Width="120" ></asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#4d81b3" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#DCDCDC" />
        </asp:GridView>
        </div>
        </td>
    </tr>
    </table>
</asp:Content>

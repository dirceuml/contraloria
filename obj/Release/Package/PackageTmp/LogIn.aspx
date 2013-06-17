<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LogIn.aspx.vb" Inherits="ATVContraloria.LogIn" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>ATV Contraloria</title>
<style type="text/css">
<!--
body {
	background-image: url(images/fondo.png);
	background-repeat: no-repeat;
	background-position: left top;
	margin: 0px;
	padding: 0px;
}
#tablaform {
	float: right;
	padding-top: 0px;
	padding-right: 30px;
	padding-bottom: 0px;
	padding-left: 0px;
}
#tablecontainer {
	display: block;
	float: right;
}
#formcontainer img {
	display: block;
	float: left;
	margin: 0px;
	padding: 0px;
}
-->
</style>
<link href="css/estilos.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
      <div id="tablewrap">
        <div id="formcontainer">
        <img src="images/logo.png" width="724" height="157" />
        <div id="tablecontainer"><div id="tablaform">
          <table width="100%" border="0" id="tabla">
            <tr>
              <td class="titulologin">Cód. Usuario:</td>
              <td><asp:TextBox ID="txtCodUsuario" runat="server" class="textoxlogin"></asp:TextBox></td>
            </tr>
            <tr>
              <td class="titulologin">Contraseña:</td>
              <td><asp:TextBox ID="txtPassword" runat="server" TextMode="Password" class="textoxlogin"></asp:TextBox></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><div id="botonlogin"><asp:LinkButton ID="lnkIngresar" runat="server"><span>Ingresar</span></asp:LinkButton></div></td>
            </tr>
          </table>
            <asp:Label ID="lblMensaje" runat="server" Text="" BackColor="White" ForeColor="Red" Font-Bold="true"></asp:Label>
        </div>
      </div>
      </div>
      </div>
    </form>
</body>
</html>

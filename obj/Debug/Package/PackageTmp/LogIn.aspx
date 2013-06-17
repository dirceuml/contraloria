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
              <td class="titulologin">Usuario:</td>
              <td><input name="textfield" type="text" class="textoxlogin" id="textfield" /></td>
            </tr>
            <tr>
              <td class="titulologin">Contraseña:</td>
              <td><input name="textfield2" type="text" class="textoxlogin" id="textfield2" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><div id="botonlogin"><a href="ReportesFC.aspx"><span>Ingresar</span></a></div></td>
            </tr>
          </table>
        </div>
      </div>
    </form>
</body>
</html>

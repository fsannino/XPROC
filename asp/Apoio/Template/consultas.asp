<%
opti=request("op")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Apoiadores</title>
<base target="principal">
<style type="text/css">
<!--
.style1 {font-size: xx-small}
.style2 {
	font-size: small;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #000063;
}
-->
</style>
</head>

<body>

<h6 align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0; font-size: small;"><img src="menu5.gif" width="170" height="360" border="0" usemap="#FPMap0Map"> </h6>
<p align="center" class="style2">Coordenadores</p>
<p align="center" class="style1"><img src="../Clis/Template/menu_cli.gif" width="158" height="85" border="0" usemap="#FPMap0MapMap">
    <map name="FPMap0MapMap">
        <area shape="rect" coords="13,48,151,66" href="../Clis/Template/por_apoio.asp?op=<%=opti%>" alt="Consulta por &Oacute;rg&atilde;o">
        <area shape="rect" coords="14,19,150,38" href="../Clis/Template/por_nome.asp?op=<%=opti%>" alt="Consulta Ordenada por Nome">
      </map>
    <map name="FPMap0Map">
        <area shape="rect" coords="27,194,165,212" href="por_apoio.asp?op=<%=opti%>" alt="Consulta por &Oacute;rg&atilde;o Apoiado">
        <area shape="rect" coords="27,78,165,97" href="por_modulo.asp?op=<%=opti%>" alt="Consulta por M&oacute;dulo">
        <area shape="rect" coords="27,54,164,73" href="por_lotacao.asp?op=<%=opti%>" alt="Consulta por &Oacute;rg&atilde;o de Lota&ccedil;&atilde;o">
        <area shape="rect" coords="28,32,164,51" href="por_nome.asp?op=<%=opti%>" alt="Consulta Ordenada por Nome">
        <area shape="rect" coords="27,216,165,234" href="por_apoio_modulo.asp" alt="Consulta por &Oacute;rg&atilde;o Apoiado X M&oacute;dulo">
        <area shape="rect" coords="26,238,164,256" href="modulo_n_apoiado.asp" alt="Consulta por &Oacute;rg&atilde;o X M&oacute;dulos n&atilde;o Apoiados">
        <area shape="rect" coords="27,310,160,330" href="multiplicador.asp?op=<%=opti%>" alt="Rela&ccedil;&atilde;o de Multiplicadores">
        <area shape="rect" coords="27,101,164,119" href="por_momento.asp?op=<%=opti%>" alt="Consulta por Momento">
        <area shape="rect" coords="26,147,163,165" href="por_onda.asp?op=<%=opti%>">
        <area shape="rect" coords="28,335,159,354" href="multiplicador_modulo.asp?op=<%=opti%>">
        <area shape="rect" coords="27,260,164,279" href="obs_modulo.asp?op=<%=opti%>">
        <area shape="rect" coords="26,122,164,141" href="por_momento_modulo.asp?op=<%=opti%>">
        <area shape="rect" coords="26,170,164,188" href="por_onda_modulo.asp?op=<%=opti%>">
      </map>
</p>
<table width="95%" border="0" style="margin-bottom: 0">
  <tr> 
    <td width="13%"><div align="right"><a href="javascript:history.go(-1)" target="_top"><img border="0" src="../volta_f02.gif"></a></div></td>
    <td width="87%"><div align="left"><b><font face="Verdana" color="#000080" size="1">Retornar</font></b></div></td>
  </tr>
  <tr> 
    <td><div align="right"><a href="selecione.htm"><img src="../excel.jpg" width="25" height="21" border="0"></a></div></td>
    <td><div align="left"><b><font face="Verdana" color="#000080" size="1">Apoiadores para o 
      Excel</font></b> </div></td>
  </tr>
  <tr>
    <td height="26"><div align="right"><a href="../Clis/Template/por_nome.asp?op=<%=opti%>&excel=1" target="_blank"><img src="../excel.jpg" width="25" height="21" border="0"></a></div></td>
    <td><div align="left"><b><font face="Verdana" color="#000080" size="1">Coordenadores para o Excel</font></b> </div></td>
  </tr>
</table>

</body>

</html>

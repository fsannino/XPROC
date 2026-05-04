<%
chave=ucase(REQUEST("CHAVE"))

if chave="" then
	response.redirect "erro.asp"	
end if
%>
<HEAD>
<title>Selecione a Aplicação desejada</title>
<style>
a {text-decoration:none;}
a:hover {
	text-decoration:underline;
	color: #006666;
}
a:link {
	color: #006666;
}
a:visited {
	color: #006666;
}
a:active {
	color: #006666;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></HEAD>
<body bgcolor="#FFFFFF" text="#000000">
 <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
 <p style="margin-top: 0; margin-bottom: 0" align="center">
 <font face="Verdana" size="5">
<strong>Selecione a Aplica&ccedil;&atilde;o Desejada </strong></font></p>
 <p align="center">&nbsp;</p>
 <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana"><img src="../../imagens/b011.gif" width="16" height="16"> <a href="index2.asp?chave=<%=chave%>">Apoiadores Locais / Multiplicadores</a></font></p>
 <p style="margin-top: 0; margin-bottom: 0" align="center"><BR>
 <p align="center">   <font face="Verdana"> <img src="../../imagens/b011.gif" width="16" height="16"> <a href="clis/index.asp?chave=<%=chave%>">Coordenadores Locais de Implanta&ccedil;&atilde;o </a></font>
 <p style="margin-top: 0; margin-bottom: 0" align="center">  </p>
</body>
</html>
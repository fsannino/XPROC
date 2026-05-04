<%
Acesso = Session("Acesso")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>#BACKLOG - Solicitações de Melhoria no SAP R/3#</title>
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body bgcolor="#FFFFFF" link="#000099" vlink="#000099" alink="#000099">
<div align="center">
  <table width="73%" border="0" height="437">
    <tr> 
      <td height="346"> 
        <div align="center"><img src="logo.jpg" width="716" height="336"></div>
      </td>
    </tr>
    <%
	if acesso=1 then
	%>
    <tr> 
      <td height="38"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000066"><b><font size="3"><a href="cad_backlog.asp">Cadastro 
          de Solicita&ccedil;&atilde;o</a></font></b></font></div>
      </td>
    </tr>
    <%
	end if	
	%>
    <tr> 
      <td height="38"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000066"><b><font size="3"><a href="tipo_consulta.asp">Consulta 
          de Solicita&ccedil;&atilde;o</a></font></b></font></div>
      </td>
    </tr>
  </table>
  <p><font face="Arial, Helvetica, sans-serif" size="1"><b><font color="#993300">Copyright 
    Gest&atilde;o do Conhecimento - Projeto Sinergia</font></b></font></p>
</div>
</body>

</html>
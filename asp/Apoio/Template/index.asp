<%
opti=request("op")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Base de Apoiadores Locais</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset framespacing="0" border="0" cols="251,*" frameborder="0">
  <%if opti<>0 then%>
  <frame name="Conteudo" target="principal" src="menu.asp?op=<%=opti%>">
  <%else%>
  <frame name="Conteudo" target="principal" src="consultas.asp?op=<%=opti%>"> 
  <%end if%> 
  <frame name="principal" src="corpo.asp" scrolling="auto">
  <noframes>
  <body>

  <p>Esta página usa quadros mas seu navegador não aceita quadros.</p>

  </body>
  </noframes>
</frameset>

</html>

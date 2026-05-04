<%
Session("Tabela") = "SINERGIA"

if Session("Data_inicio") = "" then
	Session("Data_inicio") = "01/09/2003"
end if

if Session("Periodo")="" or Session("Periodo")=0 then
	Session("Periodo") = 7
end if

if Session("Erro")="" or Session("Erro")="TODOS" then
	Session("Erro")="TODOS"
	Session("Compl")=""
end if

if Session("Orgao")="" or Session("Orgao")="TODOS" then
	Session("Orgao")="TODOS"
	Session("Compl") = Session("Compl")
end if

if Session("Modo") = "" then
	Session("Modo") = "Q"
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Acompanhamento de Chamados</title>
</head>

<frameset cols="252,*" framespacing="0" border="0" frameborder="0">
          <frame name="Conteudo" target="principal" src="menu.asp" scrolling="auto" noresize>
          <frame name="principal" src="target.asp" scrolling="auto" noresize>
          <noframes>

          <body>

          <p>Esta página usa quadros mas seu navegador não aceita quadros.</p>
          </body>

          </noframes>
</frameset>

</html>
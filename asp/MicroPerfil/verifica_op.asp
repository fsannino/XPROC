<%
opcao = request("opt")
micro = request("selMicro")

if opcao = 2 then
	response.redirect "altera_micro.asp?selMicroPerfil=" & micro
else
	if opcao = 3 then
			response.redirect "exclui_micro.asp?selMicro=" & micro
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>
<body>
</body>
</html>

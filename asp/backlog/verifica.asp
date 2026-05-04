<!--#include file="conecta.asp" -->
<%
set db = server.createobject("ADODB.CONNECTION")

db.open Session("Conn_String_Cogest_Gravacao")

set rs = db.execute ("SELECT * FROM BACKLOG_CHAVE WHERE USMA_CD_USUARIO='" & request("CHAVE") & "'")

if rs.eof=false then
		Session("Acesso")=1
		response.redirect "menu.asp"			
	else
		Session("Acesso")=0			
		response.redirect "menu.asp"
end if

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Redirecionando...</title>
</head>

<body>
</body>

</html>
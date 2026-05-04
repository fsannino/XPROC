<!--#include file="conecta.asp" -->
<%
set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

set rs = db.execute ("SELECT * FROM CLI WHERE USMA_CD_USUARIO='" & request("CHAVE") & "'")

if rs.eof=false or request("chave")="SD02" or request("chave")="RV61" or  request("chave")="XT41" or request("chave")="PE10" or request("chave")="XK79" or request("chave")="XD47" or request("chave")="SM23" or request("chave")="DCX0" or request("chave")="B511" or request("chave")="EADE" or request("chave")="WS04" then

	if request("chave")="XD47" then

		Session("Acesso")=1
		response.redirect "menu.asp"			

	else

		
		if request("chave")="XT41" or request("chave")="SD02" or request("chave")="RV61" or request("chave")="PE10" or request("chave")="XK79" then
			Session("Acesso")=1		
			response.redirect "consulta.asp"		
		else
			Session("Acesso")=0			
			response.redirect "menu.asp"
		end if

	end if

else
			Session("Acesso")=1		
			response.redirect "consulta.asp"		
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
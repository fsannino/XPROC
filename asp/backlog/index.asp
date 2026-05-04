<title>Validação de Usuário</title>
<%
response.buffer = true

set objUSR = server.createobject("Seseg.Usuario")

if objUSR.GetUsuario then
    response.clear
 	response.redirect "verifica.asp?chave=XD47" & objUSR.seiChave
	response.end
end if
%>
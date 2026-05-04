<%
if request("op")=1 then
	mensagem="USUÁRIO NÃO CADASTRADO NO SISTEMA"
end if

if request("op")=2 then
	mensagem="USUÁRIO NÃO POSSUI PERMISSÃO PARA ACESSAR O SISTEMA"
end if

if request("op")=3 then
	mensagem="AS INFORMAÇÕES DE SUA SESSÃO FORAM PERDIDAS. POR FAVOR, REINICIE O APLICATIVO..."
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body topmargin="0" leftmargin="0">
<form>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="487">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="418" valign="top"><img border="0" src="lado.jpg" width="83" height="445"></td>
                      <td width="87%" height="418" valign="top">
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;<p>&nbsp;<p align="center"><b><font face="Verdana" size="2" color="#800000"><%=mensagem%></font></b><p>&nbsp;</td>
           </tr>
</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>
<%
str_Cd_Erro = request("pCdMsgErro")
if str_Cd_Erro = 1 then
	msg_erro = "Usuário não encontrado."
elseif str_Cd_Erro = 2 then
	msg_erro = "Senha não é válida."
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>
<script language="JavaScript" type="text/JavaScript">
    opener.opener = opener;
    opener.close();
</script>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:jump_58(); window.close()">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">&nbsp;    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="85%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="70%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="70%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="70%" height="36"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b></b></font></td>
  </tr>
  <tr> 
    <td width="70%" height="24"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><center><%=msg_erro%></center></font></td>
  </tr>
  <tr> 
    <td width="70%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="70%">&nbsp;</td>
  </tr>
  <tr>
    <td width="70%"><div align="center"><strong><a href="sinergialogin.html">Tentar Novamente</a></strong>  / <a href="javascript:window.close()"><strong>Sair</strong></a></div></td>
  </tr>
  <tr>
    <td width="70%"> 
      <div align="right"></div>
    </td>
  </tr>
  <tr>
    <td width="70%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>

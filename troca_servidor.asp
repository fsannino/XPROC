
<%
str_Anterior = Session("Conn_String_Cogest_Gravacao")
'RESPONSE.Write(Session("Conn_String_Cogest_Gravacao"))
if Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest" then
   Session("Conn_String_Cogest_Gravacao")     = "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
   str_Atual = Session("Conn_String_Cogest_Gravacao")
elseif Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest" THEN
   Session("Conn_String_Cogest_Gravacao")     = "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest" 
   str_Atual = Session("Conn_String_Cogest_Gravacao")
end if

' onload="javascript:window.close()"

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<table width="98%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="imagens/voltar.gif"></a> 
            </div></td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="imagens/avancar.gif"></a></div></td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="imagens/favoritos.gif"></a></div></td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="imagens/imprimir.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="imagens/atualizar.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <p align="center"><a href="indexA.asp"><img src="imagens/home.gif" width="19" height="20" border="0"></a> 
          </td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="83%">&nbsp;</td>
    <td width="8%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><font color="#330099" size="3" face="Verdana, Arial, Helvetica, sans-serif">Troca 
      de Servidor: PRODU&Ccedil;&Atilde;O &lt;-&gt; TREINAMENTO</font></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Anterior:</font></div></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Anterior%></font></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Atual:</font></div></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Atual%></font></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><div align="center"><a href="troca_servidor.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Troca 
        servidor</font></a></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><div align="center"></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>

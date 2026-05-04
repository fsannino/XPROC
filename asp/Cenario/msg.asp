<%
'if request("pOpt") <> "" then
   str_Opt = request("pOpt")
'else
'   str_Opt = ""
'end if

'response.write str_Opt
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="88%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="96%" height="36"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000099"> 
      <% if str_Opt = 0 then
	      str_Msg = "Não encontrado o cenário : "   & request("txtCenario")
		  int_Volta_Tela = "-2"
         elseif str_Opt = 1 then
	      str_Msg = ""
		  str_funcao = ""	   
		  int_Volta_Tela = "-3"
         elseif str_Opt = 2 then
	      str_Msg = "  "
		  str_funcao = ""	   
		  int_Volta_Tela = "-3"
         elseif str_Opt = 3 then
	      str_Msg = "  "
		  str_funcao = ""	   
		  int_Volta_Tela = "-3"
         elseif str_Opt = 4 then
	      str_Msg = "  "
		  int_Volta_Tela = "-2"
	   end if 
	%>
      </font></b></font></td>
  </tr>
  <tr> 
    <td width="96%" height="24"> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000099"><%=str_Msg%></font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="96%"><font color="#000099"><b></b></font></td>
  </tr>
  <tr> 
    <td width="96%"> 
      <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3">&nbsp; 
        </font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="96%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="96%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="96%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="96%"> 
      <div align="right"> </div>
    </td>
  </tr>
  <tr> 
    <td width="96%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="278">&nbsp;</td>
          <td height="1" valign="middle" align="left" width="24"> 
            <div align="right"><a href="../../indexA.asp"> <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></div>
          </td>
          <td height="30" valign="middle" align="left" width="357"> <font face="Verdana" color="#330099" size="2">Retornar 
            para Tela Principal</font></td>
        </tr>
        <% if str_Opt = 0 OR str_Opt = 1  then %>
        <tr> 
          <td width="278">&nbsp;</td>
          <td height="1" valign="middle" align="left" width="24"> 
            <div align="right"><a href="JavaScript:history.back(<%=int_Volta_Tela%>)"> 
              <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></div>
          </td>
          <td height="30" valign="middle" align="left" width="357"> <font face="Verdana" color="#330099" size="2">Retornar 
            para Tela Sele&ccedil;&atilde;o de Cen&aacute;rio</font></td>
        </tr>
        <% elseif str_Opt = 4 then %>
        <tr> 
          <td width="278">&nbsp;</td>
          <td height="1" valign="middle" align="left" width="24"> 
            <div align="right"><a href="../MacroPerfil/seleciona_macro_perfil.asp?pOPT=1"> <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></div>
          </td>
          <td height="30" valign="middle" align="left" width="357"> <font face="Verdana" color="#330099" size="2">Retornar 
            para Tela - n&atilde;o usado</font></td>
        </tr>
        <% elseif str_Opt = 2 then %>
        <tr> 
          <td width="278">&nbsp;</td>
          <td height="1" valign="middle" align="left" width="24"> 
            <div align="right"><a href="../MacroPerfil/seleciona_macro_perfil.asp?pOPT=2"> <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></div>
          </td>
          <td height="30" valign="middle" align="left" width="357"> <font face="Verdana" color="#330099" size="2">Retornar 
            para Tela - n&atilde;o usado</font></td>
        </tr>
        <% end if %>
        <tr> 
          <td width="278"><font color="#CCCCCC"><%=str_Del%> </font></td>
          <td width="24"> 
            <div align="right"></div>
          </td>
          <td width="357">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>

<%@LANGUAGE="VBSCRIPT"%> 
<%

strMSG = Request("pMsg")
'response.Write(strMSG)
if strMSG = "C" then
   str_Tipo = "C"
   strMSG = "Acesso apenas para consulta"
end if
strPlano =  Request("pPlano")
strUsuario = Request("pUsua")

strErroServidor = Request("pErroServidor")

%>
<html>
<head>
<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href.href='altera_Atividade1.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }
</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"><table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr>
          <td bgcolor="#330099" width="39" valign="middle" align="center">
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>       
          </div></td>
          <td bgcolor="#330099" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div></td>
          <td bgcolor="#330099" width="27" valign="middle" align="center">
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div></td>
        </tr>
        <tr>
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="627" height="150" border="0" cellpadding="5" cellspacing="5">
    <tr>
      <td height="29"></td>
      <td height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="29"></td>
      <td height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="101" height="29"></td>
      <td width="125" height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;        </td>
    </tr>
    <%else%>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2">&nbsp; </td>
    </tr>
    <%end if%>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="51"><a href="../../indexA.asp"><img src="../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
      <td height="1" valign="middle" align="left" width="458"><font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
    </tr>
    <%if strUsuario <> "" then%>
    <%else
	       if str_Tipo <> "C" and strErroServidor = "" then%>
    <%   end if
	   end if%>
    <%if strPlano <> "" and strErroServidor = "" then%>
    <%end if%>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
  </table>
</form>
</body>
</html>

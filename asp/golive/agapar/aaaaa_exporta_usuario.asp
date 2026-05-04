<%@LANGUAGE="VBSCRIPT"%> 
<%



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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../../imagens/voltar.gif"></a>       
          </div></td>
          <td bgcolor="#330099" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../../imagens/avancar.gif"></a></div></td>
          <td bgcolor="#330099" width="27" valign="middle" align="center">
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../../imagens/favoritos.gif"></a></div></td>
        </tr>
        <tr>
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
            <div align="center"><a href="javascript:print()"><img border="0" src="../../../imagens/imprimir.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../../imagens/atualizar.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
            <div align="center"><a href="../../../indexA.asp"><img src="../../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
</form>
</body>
</html>

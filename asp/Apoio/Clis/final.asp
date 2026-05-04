<%
select case request("edita")
	case 0
		topico="O Registro foi incluído com sucesso!"
	case 1
		topico="O Registro foi Alterado com sucesso!"
end select		
%>
<html>
<head>
<title>Base de Dados de Coordenadores Locais</title>
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
}
</script>

<body topmargin="0" leftmargin="0" onKeyDown="verifica_tecla()">
<form method="POST" action="" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"></td>
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font size="3" face="Verdana" color="#000080">Apoiadores
  Locais e Multiplicadores</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2"><%=topico%></font><font face="Verdana" color="#330099" size="3"></font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="889" height="113">
  <tr>
    <td width="287" height="38"></td>
            <td width="26" height="38"><a href="menu.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="38" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="287" height="40"></td>
            <td width="26" height="40">
              <p align="right"><a href="cad_orgao.asp?chave=<%=request("chave")%>&amp;atrib=<%=request("atrib")%>"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="40" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Associar
              Órgãos Apoiados</font></td>
  </tr>
  <tr>
    <%if request("pai")=0 then%>
    <%end if%>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>


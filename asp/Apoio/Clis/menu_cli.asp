<%@LANGUAGE="VBSCRIPT"%>
<%
if session("CdUsuario")="" then
	response.redirect "Index.asp"
end if
%>
<!--#include file="../conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

cont = session("Conn_String_Cogest_Gravacao")
select case cont
	case "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"
		tipo="P"	
	case "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
		tipo="D"	
	case "Provider=SQLOLEDB.1;server=S6000DB15;pwd=treinasin00;uid=treinasin;database=cogest"
		tipo="T"
end select


cli=request("cli")
Session("Cli") = cli

set rs=db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO ORDER BY ORME_CD_ORG_MENOR")
%>

<html>
<head>

<title>Base de Dados de Coordenadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
//alert(message);
//return false;
}
}
if (document.layers) {
if (e.which == 3) {
//alert(message);
//return false;
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

function procura(e)
{
if (e==1)
{
window.open("procura.asp?apoio=1&op=1","_blank","width=240,height=260,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
if (e==2)
{
window.open("procura.asp?apoio=1&op=2","_blank","width=240,height=260,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
}

function procura_cli(e)
{
if (e==1)
{
window.open("procura.asp?apoio=1&op=1","_blank","width=240,height=150,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
if (e==2)
{
window.open("procura.asp?apoio=1&op=2","_blank","width=240,height=150,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
if (e==3)
{
window.open("procura.asp?apoio=1&op=3","_blank","width=240,height=150,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
}

</SCRIPT>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000000" alink="#000000" onKeyDown="verifica_tecla()" link="#000000">
<form name="frm1" method="POST" action="valida_cad_orgao.asp">

  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20" colspan="2">&nbsp;</td>
      <td width="44%" height="60" colspan="2"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
      <td width="36%" valign="top" colspan="2"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> <div align="center"> 
                <p align="center"> </div></td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../itens.asp?chave=<%=Session("CdUsuario")%>"><img src="../../../imagens/voltar.gif" alt="Sele&ccedil;&atilde;o de Aplica&ccedil;&atilde;o" width="30" height="30" border="0"></a></div></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20"><font color="#003366" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=tipo%></font></td>
      <td height="20">&nbsp; </td>
      <td height="20"></td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" face="Verdana" size="4">
    COORDENADORES LOCAIS DE IMPLANTAÇÃO (CLI) </font></b></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table width="72%" border="0" align="center">
  <table width="856" height="227" border="0">
    <tr>
      <td width="182" rowspan="3" height="65" align="center"></td>
      <td width="157" rowspan="3" height="65" align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><img src="../Clis/cad_cord.jpg" width="78" height="60"></strong></font></td>
      <td width="2" height="23"></td>
      <td width="489" height="23"></td>
    </tr>
    <tr>
      <td width="2" height="23"></td>
      <td width="489" height="23"><b><font color="#000080" face="Verdana" size="2"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura_cli(1)">INCLUIR COORDENADOR LOCAL DE IMPLANTAÇÃO</a></font></strong></font></b></td>
    </tr>
    <tr>
      <td width="2" height="11"></td>
      <td width="489" height="11"></td>
    </tr>
    <tr>
      <td width="182" height="16" align="center"></td>
      <td width="157" height="16" align="right"></td>
      <td width="2" height="16"></td>
      <td width="489" height="16"></td>
    </tr>
    <tr>
      <td width="182" rowspan="3" height="38" align="center"></td>
      <td width="157" rowspan="3" height="38" align="right"><img border="0" src="alterar.jpg" width="88" height="56"></td>
      <td width="2" height="23"></td>
      <td width="489" height="23"></td>
    </tr>
    <tr>
      <td width="2" height="22"></td>
      <td width="489" height="22"><b><font face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="2"><a href="#" onClick="javascript:procura_cli(2)">ALTERAR / EXCLUIR COORDENADOR LOCAL DE IMPLANTAÇÃO</a></font></strong></font></b></td>
    </tr>
    <tr>
      <td width="2" height="1"></td>
      <td width="489" height="1"></td>
    </tr>
    <tr>
      <td width="182" height="13" align="center"></td>
      <td width="157" height="13" align="right"></td>
      <td width="2" height="1"></td>
      <td width="489" height="1"></td>
    </tr>
    <tr>
      <td width="182" rowspan="3" height="70" align="center"></td>
      <td width="157" rowspan="3" height="70" align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><img src="../Clis/lupa2.jpg" width="65" height="51"></strong></font></td>
      <td width="2" height="26"></td>
      <td width="489" height="26"></td>
    </tr>
    <tr>
      <td width="2" height="25"></td>
      <td width="489" height="25"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="Template/index.asp">TEMPLATES DE CONSULTAS</a></font></strong></font></b></td>
    </tr>
    <tr>
      <td width="2" height="11"></td>
      <td width="489" height="11"></td>
    </tr>
  </table>
  <p>&nbsp;</p></form>
</body>
</html>
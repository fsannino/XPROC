<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO ORDER BY ORME_CD_ORG_MENOR")
%>
<html>
<head>

<title>Base de Dados de Apoiadores Locais</title>
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

</SCRIPT>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000000" alink="#000000" onKeyDown="verifica_tecla()" link="#000000">
<form name="frm1" method="POST" action="valida_cad_orgao.asp">

  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20" colspan="2">&nbsp;</td>
      <td width="44%" height="60" colspan="2"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Session("Conn_String_Cogest_Gravacao")%></font></div></td>
      <td width="36%" valign="top" colspan="2"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center">&nbsp; 
              </div>
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
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp;
        
        </td>
      <td height="20"></td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" face="Verdana" size="4">BASE 
    DE APOIADORES LOCAIS / MULTIPLICADORES / COORDENADORES</font></b></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table width="82%" border="0" align="center">
    <tr> 
      <td colspan="2" bgcolor="#D6D6D6"> <div align="center"><font color="#000099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Apoiadores 
          Locais</strong></font></div></td>
      <td width="1%">&nbsp;</td>
      <td colspan="2" bgcolor="#CCCCCC"> <div align="center"><font color="#000099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          Multiplicadores</strong></font></div></td>
      <td width="1%">&nbsp;</td>
      <td colspan="2" bgcolor="#D6D6D6"> <div align="center"><font color="#000099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Coordenadores</strong></font></div></td>
    </tr>
    <tr> 
      <td width="8%">&nbsp;</td>
      <td width="23%">&nbsp;</td>
      <td>&nbsp;</td>
      <td width="9%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td>&nbsp;</td>
      <td width="9%">&nbsp;</td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr> 
      <td><div align="center"><img border="0" src="cadastro.jpg" width="59" height="46"></div></td>
      <td><b><font color="#000000" face="Verdana" size="1">APOIADOR LOCAL</font></b></td>
      <td>&nbsp;</td>
      <td><img border="0" src="cadastro.jpg" width="59" height="46"></td>
      <td><b><font color="#000080" face="Verdana" size="1"><a href="consulta_apoio.asp?op=1">MULTIPLICADOR</a></font></b></td>
      <td>&nbsp;</td>
      <td><div align="center"><img src="Clis/cad_cord.jpg" width="63" height="46"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">COORDENADOR</font></strong></td>
    </tr>
    <tr> 
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(1)">INCLUIR</A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(1)">INCLUIR</A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">INCLUIR</font></strong></td>
    </tr>
    <tr> 
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">ALTERAR</a></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">ALTERAR</a></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ALTERAR</font></strong></td>
    </tr>
    <tr> 
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">EXCLUIR</A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">EXCLUIR</A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">EXCLUIR&nbsp;</font></strong></td>
    </tr>
    <tr> 
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">&Oacute;RG&Atilde;OS 
        APOIADOS </A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">&Oacute;RG&Atilde;OS 
        APOIADOS </A></font></strong></td>
      <td>&nbsp;</td>
      <td><div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">&Oacute;RG&Atilde;OS 
        RELACIONADOS</A>&nbsp;</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td> <div align="right"><img src="../../imagens/b011.gif" width="16" height="16"></div></td>
      <td><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="javascript:procura(2)">CURSOS 
        APOIADOS </A></font></strong></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td height="21">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="62"> <div align="center"><img border="0" src="consulta.jpg" width="61" height="46"></div></td>
      <td><b><font color="#000080" face="Verdana" size="1"><a href="consulta_apoio.asp?op=1">VISUALIZAR 
        APOIADOR LOCAL</a></font></b></td>
      <td>&nbsp;</td>
      <td><div align="center"><img border="0" src="consulta.jpg" width="61" height="46"></div></td>
      <td><b><font color="#000080" face="Verdana" size="1"><a href="consulta_apoio.asp?op=1">VISUALIZAR 
        MULTIPLICADOR </a></font></b></td>
      <td>&nbsp;</td>
      <td><div align="center"><img src="Clis/lupa2.jpg" width="67" height="51"></div></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>VISUALIZAR 
        COORDENADOR</strong></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><div align="center"></div></td>
      <td><b><font face="Verdana" size="1"></font></b></td>
      <td>&nbsp;</td>
      <td><img border="0" src="lupa.jpg" width="55" height="45"></td>
      <td><p><b><font face="Verdana" size="1"><a href="Template/index.asp?op=1">TEMPLATES 
          DE CONSULTAS - APOIADOR LOCAL / MULTIPLICADOR</a></font></b></p></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"></div></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0">&nbsp;</p>
</form>
</body>
</html>

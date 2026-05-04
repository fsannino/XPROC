<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function Confirma()
{
if(document.frm1.txtcoment.value=="")
{
alert("Você deve espeficar seus COMENTÁRIOS");
document.frm1.txtcoment.focus();
return;
}
else
{
document.frm1.submit();
}
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_fecha_escopo.asp">
      <table width="812" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="140" height="20" colspan="2">&nbsp;</td>
      <td width="1065" height="60" colspan="3"><font color="#FFFFFF"><%'=Session("Conn_String_Cogest_Gravacao")%></font></td>
      <td width="1" valign="top"> 
        <table width="153" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="43" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="34" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="43" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="34" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="2">&nbsp; </td>
      <td height="20" width="136">&nbsp;</td>
      <td> 
        <p align="center"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma()" align="absmiddle">
      </td>
      <td height="20" width="498"> 
        <b><font face="Verdana" size="2" color="#330099">Enviar</font></b>
      </td>
      <td height="20" width="264" colspan="2">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="31%">&nbsp;</td>
      <td width="43%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="31%">&nbsp;</td>
      <td width="43%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Fechamento
        de Escopo&nbsp;</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">Comentários</font></b></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p> 
          <textarea rows="6" name="txtcoment" cols="59"></textarea>
        </p>
        </font></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function Confirma()
{
if(document.frm1.txtcoment.value=="")
{
alert("Você deve espeficar seus COMENTÁRIOS");
document.frm1.txtcoment.focus();
return;
}
else
{
document.frm1.submit();
}
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_fecha_escopo.asp">
      <table width="812" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="140" height="20" colspan="2">&nbsp;</td>
      <td width="1065" height="60" colspan="3"><font color="#FFFFFF"><%'=Session("Conn_String_Cogest_Gravacao")%></font></td>
      <td width="1" valign="top"> 
        <table width="153" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="43" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="34" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="43" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="34" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="2">&nbsp; </td>
      <td height="20" width="136">&nbsp;</td>
      <td> 
        <p align="center"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma()" align="absmiddle">
      </td>
      <td height="20" width="498"> 
        <b><font face="Verdana" size="2" color="#330099">Enviar</font></b>
      </td>
      <td height="20" width="264" colspan="2">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="31%">&nbsp;</td>
      <td width="43%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="31%">&nbsp;</td>
      <td width="43%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Fechamento
        de Escopo&nbsp;</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">Comentários</font></b></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p> 
          <textarea rows="6" name="txtcoment" cols="59"></textarea>
        </p>
        </font></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e

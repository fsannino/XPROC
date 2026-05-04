 
<!--#include file="../../asp/protege/protege.asp" -->
<%

str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

if str_MegaProcesso <> "0" then
   Session("MegaProcesso") = str_MegaProcesso
else
    if Session("MegaProcesso") <> "" then
       str_MegaProcesso = Session("MegaProcesso") 
	end if   
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQl = " Select "
str_SQL = str_SQL & " CURSO.MEPR_CD_MEGA_PROCESSO, "
str_SQL = str_SQL & " MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL = str_SQL & " CURSO.CURS_CD_CURSO, CURSO.CURS_TX_NOME_CURSO, "
str_SQL = str_SQL & " CURSO.CURS_NUM_CARGA_CURSO, "
str_SQL = str_SQL & " CURSO.CURS_TX_METODO_CURSO, "
str_SQL = str_SQL & " CURSO.CURS_TX_STATUS_CURSO, "
str_SQL = str_SQL & " CURSO.CURS_TX_DATA_TERMINO, "
str_SQL = str_SQL & " CURSO.CURS_TX_TUTOR_CURSO, "
str_SQL = str_SQL & " CURSO.CURS_TX_PUBLICO_ALVO, "
str_SQL = str_SQL & " CURSO.CURS_TX_PRE_REQUISITOS, "
str_SQL = str_SQL & " CURSO.CURS_TX_CONTEUDO_PROGRAM, "
str_SQL = str_SQL & " CURSO.CURS_TX_OBJETIVO, CURSO.CURS_TX_OBS "
str_SQL = str_SQL & " FROM CURSO INNER JOIN "
str_SQL = str_SQL & " MEGA_PROCESSO ON "
str_SQL = str_SQL & " CURSO.MEPR_CD_MEGA_PROCESSO = MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
if str_CdMegaProcesso <> 
str_SQL = str_SQL & " WHERE (CURSO.MEPR_CD_MEGA_PROCESSO = " & str_CdMegaProcesso
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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana" color="#330099" size="3">Relação 
        de Cursos por Mega-Processo</font></div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
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
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>

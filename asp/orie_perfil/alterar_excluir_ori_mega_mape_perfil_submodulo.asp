<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opt = request("txtOpt")
if request("selMegaProcesso") <> "0" then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = "0"
end if
if request("selSubModulo") <> "0" then
   str_SubModulo = request("selSubModulo")
else
   str_SubModulo = "0"
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " Select ORTE_TX_DESCRICAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_MEGA_MODULO "
str_SQL = str_SQL & " WHERE SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
str_SQL = str_SQL & " AND MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
set rdsDescModulo = db.Execute(str_SQL)
if rdsDescModulo.EOF then
   response.redirect "msg.asp?pOpt=1"
else
   str_DescSubModulo = rdsDescModulo("ORTE_TX_DESCRICAO")
end if

If str_Opt = "EM" then
   response.redirect "valida_inc_alt_exc_mega_perfil_submodulo.asp?txtOpt=EM&txtMegaProcesso="&str_MegaProcesso&"&txtSubModulo="&str_SubModulo
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT  "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESCRICAO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
set rdsMegaProcesso = db.Execute(str_SQL_MegaProc)
if not rdsMegaProcesso.EOF then
   str_DsMegaProcesso = rdsMegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
else
   str_DsMegaProcesso = "NÃO ENCONTRADO O MEGA"
end if
rdsMegaProcesso.close
set rdsMegaProcesso = nothing

SQL_SubModulo=""
SQL_SubModulo = SQL_SubModulo & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_SubModulo = SQL_SubModulo & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_SubModulo = SQL_SubModulo & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_SubModulo = SQL_SubModulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
SQL_SubModulo = SQL_SubModulo + " WHERE SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
set rdsSubModulo = db.Execute(SQL_SubModulo)
if not rdsSubModulo.EOF then
   str_DsSubModulo = rdsSubModulo("SUMO_TX_DESC_SUB_MODULO")
else
   str_DsSubModulo = "NÃO ENCONTRADO O SUB"
end if
rdsSubModulo.close
set rdsSubModulo = nothing

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " PERFIL_ORIEN_MEGA_MODULO.MEPR_CD_MEGA_PROCESSO "
str_SQL = str_SQL & " , PERFIL_ORIEN_MEGA_MODULO.SUMO_NR_CD_SEQUENCIA "
str_SQL = str_SQL & " , PERFIL_ORIEN_MEGA_MODULO.ORTE_TX_DESCRICAO "
str_SQL = str_SQL & " , SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
str_SQL = str_SQL & " FROM dbo.PERFIL_ORIEN_MEGA_MODULO INNER JOIN dbo.SUB_MODULO ON dbo.PERFIL_ORIEN_MEGA_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
str_SQL = str_SQL & " where  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL = str_SQL & " order by SUMO_TX_DESC_SUB_MODULO "
set rdsOrient = db.Execute(str_SQL)

%>
<html>
<head>
<script>
function manda()
{
//alert('altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value)
//+'&selSubModulo='+
//alert(document.frm1.selSubModulo.value)
window.location.href='altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value
}

function Confirma()
{
  if (document.frm1.txtOrientacoes1.value == "")
     { 
	 alert("O preenchimento do campo orientações é obrigatório!");
     document.frm1.txtOrientacoes1.focus();
     return;
     }
  else 
     {
     document.frm1.submit(); 
	 }
}
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_inc_alt_exc_mega_perfil_submodulo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20"><table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
            <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
            <td width="195"></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28"></td>
            <td width="26">&nbsp;</td>
            <td width="159"></td>
          </tr>
        </table> </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="53%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="53%"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alterar</font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
          descri&ccedil;&atilde;o para Assuntos - MEGA</font></div></td>
      <td width="26%"><a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>">Relat&oacute;rio 
        completo</a> </font></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=str_DsMegaProcesso%></strong></font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
        <input name="txtOpt" type="hidden" value="AM">
        <input name="txtMegaProcesso" type="hidden" id="txtMegaProcesso" value="<%=str_MegaProcesso%>">
        <input name="txtSubModulo" type="hidden" id="txtSubModulo" value="<%=str_SubModulo%>">
        </strong></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Assuntos 
        - <font size="3"><%=str_DsSubModulo%></font></font> </td>
      <td></td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
        - </font><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Digite 
        as descri&ccedil;&otilde;es finalizando cada par&aacute;grafo com um &lt;br&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font></td>
      <td></td>
    </tr>
    <tr> 
      <td width="6%"></td>
      <td width="88%"><textarea name="txtOrientacoes1" cols="110" rows="5" dir="ltr"><%=str_DescSubModulo%></textarea></td>
      <td width="6%"></td>
    </tr>
    <tr> 
      <td></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table> 
  <table width="89%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
    <tr bgcolor="#FFFFFF"> 
      <td height="35" colspan="2"> <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Assuntos 
          cadastrados</strong></font></div></td>
    </tr>
    <tr> 
      <td width="21%" bgcolor="#0000FF"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Assunto</font></strong></td>
      <td width="79%" bgcolor="#0000FF"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></td>
    </tr>
    <% Do While not rdsOrient.EOF %>
    <tr> 
      <td valign="top" bgcolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("SUMO_TX_DESC_SUB_MODULO")%></font></td>
      <td valign="top" bgcolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("ORTE_TX_DESCRICAO")%></font></td>
    </tr>
    <%  rdsOrient.movenext
	Loop %>
    <tr> 
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>

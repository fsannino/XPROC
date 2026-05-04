<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("selMegaProcesso") <> "0" then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = "0"
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

atual=0

set rs=db.execute("SELECT MAX(ORTE_NR_SEQUENCIAL) AS CODIGO FROM  PERFIL_ORIEN_MEGA_TERMOS")

if not isnull(rs("CODIGO")) then
	atual = rs("CODIGO")
end if

if atual=0 then
	atual=0
else
	atual=atual
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

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " ORTE_NR_SEQUENCIAL "
str_SQL = str_SQL & " , ORTE_TX_TERMO "
str_SQL = str_SQL & " , ORTE_TX_DESCRICAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_MEGA_TERMOS "
str_SQL = str_SQL & " where  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL = str_SQL & " order by ORTE_TX_TERMO "

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
   document.frm1.submit(); 
}
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_inc_alt_exc_mega_perfil_termos.asp">
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
      <td width="53%"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Incluir 
          termos novos Relevantes - Perfil - MEGA</font></div></td>
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
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%"></td>
      <td width="88%"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Voc&ecirc; 
          poder&aacute; cadastrar quantas Termos sejam necess&aacute;rias. </font></div></td>
      <td width="6%">&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">A 
          cada inclus&atilde;o poder&atilde;o ser cadastradas at&eacute; 4 orienta&ccedil;&otilde;es.</font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><input name="txtOpt" type="hidden" value="IT">
        <font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>
        <input name="txtMegaProcesso" type="hidden" id="txtMegaProcesso" value="<%=str_MegaProcesso%>">
        </strong></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Termo</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td> <input name="txtTermo1" type="text" id="txtTermo1" size="80" maxlength="150"></td>
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
      <td width="88%"><textarea name="txtOrientacoes1" cols="110" rows="5" dir="ltr"></textarea></td>
      <td width="6%"></td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">- &Eacute; 
        obrigat&oacute;rio o preenchimento dos campos &quot;Termo&quot; e &quot;Descri&ccedil;&atilde;o&quot; 
        para que a orienta&ccedil;&atilde;o seja cadastrada.</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Termo</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%"></td>
      <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="txtTermo2" type="text" id="txtTermo2" size="80" maxlength="150">
        </font></td>
      <td width="6%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p>&nbsp; </p>
        </font></td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
        - </font><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Digite 
        as descri&ccedil;&otilde;es finalizando cada par&aacute;grafo com um &lt;br&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%"></td>
      <td width="88%"><textarea name="txtOrientacoes2" cols="110" rows="5" dir="ltr"></textarea></td>
      <td width="6%">&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">- &Eacute; 
        obrigat&oacute;rio o preenchimento dos campos &quot;Termo&quot; e &quot;Descri&ccedil;&atilde;o&quot; 
        para que a orienta&ccedil;&atilde;o seja cadastrada.</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Termo</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%"></td>
      <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="txtTermo3" type="text" id="txtTermo3" size="80" maxlength="150">
        </font></td>
      <td width="6%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
        - </font><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Digite 
        as descri&ccedil;&otilde;es finalizando cada par&aacute;grafo com um &lt;br&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="88%"><textarea name="txtOrientacoes3" cols="110" rows="5" dir="ltr"></textarea></td>
      <td width="6%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">- &Eacute; 
        obrigat&oacute;rio o preenchimento dos campos &quot;Termo&quot; e &quot;Descri&ccedil;&atilde;o&quot; 
        para que a orienta&ccedil;&atilde;o seja cadastrada.</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Termo</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="txtTermo4" type="text" id="txtTermo4" size="80" maxlength="150">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
        - </font><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Digite 
        as descri&ccedil;&otilde;es finalizando cada par&aacute;grafo com um &lt;br&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><textarea name="txtOrientacoes4" cols="110" rows="5" dir="ltr"></textarea></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">- &Eacute; 
        obrigat&oacute;rio o preenchimento dos campos &quot;Termo&quot; e &quot;Descri&ccedil;&atilde;o&quot; 
        para que a orienta&ccedil;&atilde;o seja cadastrada.</font></td>
      <td>&nbsp;</td>
    </tr>
  </table> 
  <table width="89%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
    <tr bgcolor="#FFFFFF"> 
      <td height="35" colspan="3"> <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Termos 
          cadastrados</strong></font></div></td>
    </tr>
    <tr> 
      <td width="3%" bgcolor="#0000FF"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">ID</font></strong></div></td>
      <td width="11%" bgcolor="#0000FF"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Termo</font></strong></td>
      <td width="86%" bgcolor="#0000FF"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></td>
    </tr>
    <% Do While not rdsOrient.EOF %>
    <tr> 
      <td valign="top" bgcolor="#FFFFFF"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("ORTE_NR_SEQUENCIAL")%></font></div></td>
      <td valign="top" bgcolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("ORTE_TX_TERMO")%></font></td>
      <td valign="top" bgcolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("ORTE_TX_DESCRICAO")%></font></td>
    </tr>
    <%  rdsOrient.movenext
	Loop %>
    <tr> 
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>

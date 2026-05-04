<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opt = request("txtOpt")
if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if

if request("selOrient") <> "" then
   str_Cd_Orient = request("selOrient")
else
   str_Cd_Orient = "0"
end if   
'response.Write(request("selOrient"))
If str_Opt = "EO" then
   response.redirect "valida_inc_alt_exc_mega_orient_perfil.asp?txtOpt=EO&txtCdOrient="&str_Cd_Orient&"&txtMegaProcesso="&str_MegaProcesso
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL = str_SQL & " ORIE_NR_SEQUENCIAL "
str_SQL = str_SQL & " , ORIE_TX_ORIENTACOES "
str_SQL = str_SQL & " , ORIE_NR_ORDENACAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_MEGA "
str_SQL = str_SQL & " where ORIE_NR_SEQUENCIAL = " & str_Cd_Orient
str_SQL = str_SQL & " and MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
'response.Write(str_SQL)
set rdsOrient = db.Execute(str_SQL)
if not rdsOrient.EOF then
   str_Ds_Orient = rdsOrient("ORIE_TX_ORIENTACOES")
else
   str_Ds_Orient = ""
end if 
rdsOrient.close
set rdsOrient = Nothing
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
<form name="frm1" method="POST" action="valida_inc_alt_exc_mega_orient_perfil.asp">
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
          orienta&ccedil;&otilde;es ao Mapeamento</font></div></td>
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
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=str_DsMegaProcesso%></font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><input name="txtOpt" type="hidden" value="<%=str_Opt%>"> <input name="txtCdOrient" type="hidden" value="<%=str_Cd_Orient%>"> 
        <input name="txtMegaProcesso" type="hidden" value="<%=str_MegaProcesso%>"></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Digite 
        as orienta&ccedil;&otilde;es finalizando cada par&aacute;grafo com um 
        &lt;br&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%"></td>
      <td width="89%"><textarea name="txtOrientacoes1" cols="110" rows="5" dir="ltr"><%=str_Ds_Orient%></textarea></td>
      <td width="6%"></td>
    </tr>
    <tr> 
      <td width="5%"></td>
      <td width="89%">&nbsp;</td>
      <td width="6%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p>&nbsp; </p>
        </font></td>
    </tr>
  </table> 
  </form>
</body>
</html>

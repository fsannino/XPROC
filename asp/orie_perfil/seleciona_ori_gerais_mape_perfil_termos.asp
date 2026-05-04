<%@LANGUAGE="VBSCRIPT"%> 
<%

if request("pOpt")  <> "" then
   str_Opt = request("pOpt") 
else
   str_Opt = "0"
end if
if str_Opt = "A" then
   str_Titulo = "ALTERAÇÃO"
else
   str_Titulo = "EXCLUSÃO"
end if
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " ORTE_NR_SEQUENCIAL "
str_SQL = str_SQL & " , ORTE_TX_TERMO "
str_SQL = str_SQL & " , ORTE_TX_DESCRICAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_GERAL_TERMOS "
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

function Confirma2()
{
  alert(document.frm1.txtO.value) 
}

function Confirma()
{
  //alert(document.frm1.txtOpt.value) 
  if (document.frm1.selTermo.selectedIndex == 0)
     { 
	 alert("A seleção de um termo é obrigatória!");
     document.frm1.selTermo.focus();
     return;
     }      
   if(document.frm1.txtOpt.value == "A")
     {
     document.frm1.action="alterar_excluir_ori_gerais_mape_perfil_termos.asp";
     //document.frm1.target="corpo";
     document.frm1.submit();
     }
   if(document.frm1.txtOpt.value == "E")
     {
     document.frm1.action="alterar_excluir_ori_gerais_mape_perfil_termos.asp";
     //document.frm1.target="corpo";
     document.frm1.submit();
     }
}
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
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
      <td width="11%">&nbsp;</td>
      <td width="63%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="11%">&nbsp;</td>
      <td width="63%"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Seleciona 
          termos novos Relevantes - Perfil - GERAL - <%=str_Titulo%></font></div></td>
      <td width="26%"><a href="relat_ori_gerais_mape_perfil.asp"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <a href="relat_ori_gerais_mape_perfil.asp">Relat&oacute;rio completo</a> 
        </font></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td></td>
      <td><input name="txtOpt" type="hidden" value="<%=str_Opt%>"></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Selecione:</font></div></td>
      <td><select name="selTermo" size="1" id="selTermo">
          <option value="0">== Selecione um Termo ==</option>
          <% set rs=db.execute(str_SQL)
		     do until rs.eof=true %>
          <option value="<%=rs("ORTE_NR_SEQUENCIAL")%>"><%=rs("ORTE_NR_SEQUENCIAL")%> 
          - <%=Left(rs("ORTE_TX_TERMO"),90)%></option>
          <% rs.movenext
			loop %>
        </select></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="15%"></td>
      <td width="80%">&nbsp;</td>
      <td width="5%"></td>
    </tr>
    <tr> 
      <td width="15%"></td>
      <td width="80%">&nbsp;</td>
      <td width="5%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p>&nbsp; </p>
        </font></td>
    </tr>
  </table> 
  </form>
</body>
</html>

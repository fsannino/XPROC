<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Mega = request("selMegaProcesso")
str_SubModulo = request("selSubModulo") 

ssql=""
ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_INDICA_CRITICA "
ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
ssql=ssql & " WHERE "
ssql=ssql & " FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_mega 
IF str_SubModulo <> 0 THEN
   ssql=ssql & " and FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_SubModulo
END IF
ssql=ssql+"ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "


set rs=db.execute(ssql)
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<script>
function manda()
{
document.frm1.submit();
}
    
</script>
<body topmargin="0" leftmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method="POST" action="valida_func_critico.asp" name="frm1">          
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"><a href="javascript:manda()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
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

 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%"> 
        <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Associação
        de Funções Críticas</font>
      </td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
  </table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
<input type="hidden" name="TSql" size="127" value="<%=ssql%>">
<table border="0" width="85%" cellspacing="0" cellpadding="2" height="62">
  <tr>
    <td width="23%" height="21"></td>
    <td width="6%" bgcolor="#000080" height="21">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="92%" bgcolor="#000080" height="21">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="2">Função
      de Negócio</font></b></td>
  </tr>
  <%do until rs.eof=true
  if rs("FUNE_TX_INDICA_CRITICA")="1" then
  	valor="checked"
  else
  	valor=""
  end if
  
  if cor="#E4E4E4" then
  	cor="white"
  else
  	cor="#E4E4E4"
  end if
  
  %>
  <tr>
    <td width="23%" height="33"></td>
    <td width="6%" height="33" bgcolor="<%=cor%>">
      <p align="center"><font size="1"><input type="checkbox" name="func_<%=rs("fune_cd_funcao_negocio")%>" value="1" <%=valor%>></font></td>
    
    <td width="92%" height="33" bgcolor="<%=cor%>"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><a href="gera_rel_mega_funcao.asp?selMegaProcesso=<%=str_mega%>&selFuncao=<%=rs("fune_cd_funcao_negocio")%>"><%=rs("fune_cd_funcao_negocio")%></a></b> - <%=rs("fune_tx_titulo_funcao_negocio")%></font></td>
      </tr>
    <%
    rs.movenext
    loop
    %>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>

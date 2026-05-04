<%@LANGUAGE="VBSCRIPT"%>  
<%

if request("opt") = 1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
elseif request("opt") = 2 then
	Response.Buffer = TRUE
	Response.ContentType = "application/msword"
end if

str_Opc = Request("txtOpc")

str_MegaProcesso = request("selMegaProcesso")
str_Modulo=request("selSubModulo")
str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
str_Funcao=request("selFuncao")

if str_MegaProcesso <> "0" then
   str_Sql_MegaProcesso = " and FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if
if str_modulo <> "0" then
	str_Sql_SubModulo = " AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & str_modulo 
end if

if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   

if str_Uso = 1 and str_Desuso = 1 then
   str_Sql_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_Sql_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_Sql_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

if str_Funcao <> "0" then
	str_Sql_Funcao=" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao  & "'"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'response.Write("<p>" & str_Sql_MegaProcesso)
'response.Write("<p>" & str_Sql_SubModulo)
'response.Write("<p>" & str_Sql_usoDesuso )
'response.Write("<p>" & str_Sql_Funcao)


ssql=""
ssql="SELECT FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO, FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA, FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_OBS_ESPECIFICA "
ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
ssql=ssql+" where FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO > 0 " & str_Sql_MegaProcesso & str_Sql_SubModulo & str_Sql_usoDesuso & str_Sql_Funcao 
ssql=ssql+"ORDER BY FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO, FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "


'response.write str_Sql
set rs=db.execute(ssql)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
function Confirma()
{
   document.frm1.submit(); 
}

//  End -->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_ori_massa_funcao.asp">
 <input type="hidden" name="txtOpc" value="1">
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
      <td colspan="3" height="20"> <table width="625" border="0" align="center">
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
        </table></td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    Cadastro de Observa&ccedil;&atilde;o Espec&iacute;fica para Fun&ccedil;&atilde;o 
    R/3</b></font></p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 (<font face="Verdana">Clique no código da Fun&ccedil;&atilde;o R/3 para exibir
 seus dados)</font></font></font></p>
  <table width="87%" border="0">
    <tr>
      <td width="3%">&nbsp;</td>
      <td width="77%">&nbsp;</td>
      <td width="20%"><a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>">Relat&oacute;rio 
        completo</a> </font></td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>        
  <table width="87%" height="129" border="1" cellpadding="1" cellspacing="1">   
    <%
  conta=0
  DO UNTIL RS.EOF=TRUE
     str_Contador = str_Contador + 1
		%>
		<%
		str_SQL_MegaProc = ""
        str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
        str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
        str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
        str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
        str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO = " & RS("MEPR_CD_MEGA_PROCESSO")
        str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
		'response.Write(str_SQL_MegaProc)
		set rdsMegaProcesso = db.Execute(str_SQL_MegaProc)
		if not rdsMegaProcesso.EOF then
		   str_DsMegaProcesso = rdsMegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
		else
		   str_DsMegaProcesso = "NÃO ENCONTRADO"
		end if
		rdsMegaProcesso.close
		set rdsMegaProcesso = Nothing
		%>
    <tr valign="top"> 
      <td width="17%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Mega</font></strong></td>
      <td width="83%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="1"><%=str_DsMegaProcesso%> </font></strong></td>
    </tr>	
	<%
	If not isNull(RS("SUMO_NR_CD_SEQUENCIA")) then
	   str_Sql_SubModulo=""
       str_Sql_SubModulo = str_Sql_SubModulo & " SELECT SUMO_NR_CD_SEQUENCIA"
       str_Sql_SubModulo = str_Sql_SubModulo & " ,SUMO_TX_DESC_SUB_MODULO"
       str_Sql_SubModulo = str_Sql_SubModulo & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
       str_Sql_SubModulo = str_Sql_SubModulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
	   str_Sql_SubModulo = str_Sql_SubModulo + " WHERE SUMO_NR_CD_SEQUENCIA = " & RS("SUMO_NR_CD_SEQUENCIA")
	   'response.Write(str_Sql_SubModulo)	
	   set rdsSubModulo = db.Execute(str_Sql_SubModulo)
	   if not rdsSubModulo.EOF then
	      str_DsSubModulo = rdsSubModulo("SUMO_TX_DESC_SUB_MODULO")
	   else
	      str_DsSubModulo = "NÃO ENCONTRADO"
	   end if
	   rdsSubModulo.close
	   set rdsSubModulo = Nothing	
	%>		
    <tr valign="top">
      <td height="19"><strong><font face="Verdana" size="2">Assunto</font></strong></td>
      <td height="19"><font face="Verdana" size="1"><%=str_DsSubModulo%></font></td>
    </tr>
	<% end if %>
    <tr valign="top"> 
      <td height="19"><strong><font face="Verdana" size="2">Fun&ccedil;&atilde;o</font></strong></td>
      <td height="19"><strong><font face="Verdana" size="1"><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%> 
        <input name="txtFuncao<%=str_Contador%>" type="hidden" id="txtFuncao<%=str_Contador%>" value="<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>">
        </font></strong></td>
    </tr>
    <tr valign="top"> 
      <td width="17%" height="19"><font face="Verdana" size="1">&nbsp;<strong><font face="Verdana" size="2">Descrição</font></strong></font></td>
      <td width="83%" height="19"><font face="Verdana" size="1"><%=RS("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <%
		str_SQL = ""
		str_SQL = str_SQL & " SELECT dbo.ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO "
        str_SQL = str_SQL & " FROM  dbo.FUN_NEG_ORG_AGLU INNER JOIN "
        str_SQL = str_SQL & " dbo.ORGAO_AGLUTINADOR ON dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = dbo.ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "
        str_SQL = str_SQL & " WHERE (dbo.FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO = '" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "')"
		set rdsAreaAbran = db.Execute(str_SQL)
		str_AreaAbrang = ""
		if not rdsAreaAbran.EOF then
		   str_AreaAbrang = rdsAreaAbran("AGLU_SG_AGLUTINADO")
		   rdsAreaAbran.movenext
		end if   
		do while not rdsAreaAbran.EOF
		   str_AreaAbrang = str_AreaAbrang & ", " & rdsAreaAbran("AGLU_SG_AGLUTINADO")
		   rdsAreaAbran.movenext
		loop
		rdsAreaAbran.close
		set rdsAreaAbran = Nothing
		%>
    <tr valign="top"> 
      <td height="19"><strong><font face="Verdana" size="2">&Aacute;rea de Abrang&ecirc;ncia</font></strong></td>
      <td height="19"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_AreaAbrang%></font></td>
    </tr>
    <tr valign="top"> 
      <td height="19"><strong><font face="Verdana" size="2">Observa&ccedil;&atilde;o 
        Espec&iacute;fica para Fun&ccedil;&atilde;o</font></strong></td>
      <td height="19"><font face="Verdana" size="1"><strong><font face="Verdana" size="1"> 
        <textarea name="txtDescFunc<%=str_Contador%>" cols="90" rows="5"><%=RS("FUNE_TX_OBS_ESPECIFICA")%></textarea>
        <input name="txtObsAnterior<%=str_Contador%>" type="hidden" id="txtObsAnterior<%=str_Contador%>" value="<%=RS("FUNE_TX_OBS_ESPECIFICA")%>">
        </font></strong></font></td>
    </tr>
    <tr valign="top" bgcolor="#0000FF"> 
      <td height="5" bgcolor="#000099"></td>
      <td height="5" bgcolor="#000099"></td>
    </tr>
    <%
        conta=conta+1
			RS.MOVENEXT        
        	LOOP
        %>
  </table>
          <input name="txtQtdObj" type="hidden" id="txtQtdObj" value="<%=str_Contador%>">
 <P>
 <b>
 <%if conta=0 then%>
 &nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 <font face="Verdana" size="2" color="#800000">Não existe Fun&ccedil;&atilde;o R/3s para o Mega-Processo Selecionado</font>
 <%end if%>
 </b>
  <table width="75%" border="0">
    <tr>
      <td width="16%">&nbsp;</td>
      <td width="84%"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        de Fun&ccedil;&otilde;es Listadas</strong> : <%=conta%></font></td>
    </tr>
  </table>
  </form>
</body>
</html>

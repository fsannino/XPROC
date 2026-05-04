<%@LANGUAGE="VBSCRIPT"%>  
<%
server.ScriptTimeout=99999999

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")

if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   

if request("opt") = 1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
elseif request("opt") = 2 then
	Response.Buffer = TRUE
	Response.ContentType = "application/msword"
end if

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")
str_Modulo=request("selSubModulo")
'response.Write(str_Modulo)
selG=request("selG")

if str_modulo<>"0" then
	compl1=" AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & str_modulo 
end if

'response.write selG

if selG=1 then
	compl1=compl1 + " AND FUNCAO_NEGOCIO.FUNE_TX_TP_FUN_NEG ='G'"
else
   selG = 0	
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso)

if str_MegaProcesso=0 then
	IF selG=1 then
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_TX_TP_FUN_NEG ='G' " & str_usoDesuso  '& " ORDER BY MEPR_CD_MEGA_PROCESSO, FUNE_TX_TITULO_FUNCAO_NEGOCIO"
	else
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO  where MEPR_CD_MEGA_PROCESSO > 0 " & str_usoDesuso ' & " ORDER BY MEPR_CD_MEGA_PROCESSO, FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	end if
	ssql = ssql & " order by FUNE_TX_TITULO_FUNCAO_NEGOCIO "
else
	ssql=""
	ssql="SELECT distinct FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
	ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
	ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
	if request("selFuncao")<>"0" then
		ssql=ssql+" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO='" & request("selFuncao") & "' ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
	else
		ssql=ssql+" ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "	
	end if
end if

'response.write ssql

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
//  End -->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
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
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="122"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="gera_rel_geral_funcao_colunada_excel_word.asp?selMegaProcesso=<%=str_MegaProcesso%>&selSubModulo=<%=str_Modulo%>&opt=1&selG=<%=selG%>&chkEmUso=<%=str_Uso%>&chkEmDesuso=<%=str_Desuso%>&selFuncao=<%=request("selFuncao")%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></b></font></td>
            <td width="100"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="gera_rel_geral_funcao_colunada_excel_word.asp?selMegaProcesso=<%=str_MegaProcesso%>&selSubModulo=<%=str_Modulo%>&opt=2&selG=<%=selG%>&chkEmUso=<%=str_Uso%>&chkEmDesuso=<%=str_Desuso%>&selFuncao=<%=request("selFuncao")%>" target="_blank"><img src="../../imagens/exp_word.gif" width="78" height="29" border="0"></a></b></font></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Relatório
Geral de Fun&ccedil;&atilde;o R/3</b></font></p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 </font></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>        
  <table width="74%" height="62" border="1" cellpadding="1" cellspacing="1">
    <tr valign="top">
      <td width="14%"><strong><font face="Verdana" size="2">C&oacute;digo</font></strong></td> 
      <td width="16%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Fun&ccedil;&atilde;o</font></strong></td>
      <td width="24%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Descrição</font></strong></td>
      <td width="19%"><strong><font face="Verdana" size="2">&Aacute;rea de Abrang&ecirc;ncia</font></strong></td>
      <td width="27%"><strong><font face="Verdana" size="2">Observa&ccedil;&atilde;o 
        Espec&iacute;fica para Fun&ccedil;&atilde;o</font></strong></td>
    </tr>
    <%
        conta=0
        DO UNTIL RS.EOF=TRUE
		%>
    <tr valign="top">
      <td width="14%"><font face="Verdana" size="2"><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%></font></td> 
      <td width="16%" height="19"><font face="Verdana" size="1">&nbsp;<%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
      <td width="24%" height="19"><font face="Verdana" size="1"><%=RS("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
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
      <td width="19%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_AreaAbrang%></font></td>
      <%
	  set temp=db.execute("SELECT FUNE_TX_OBS_ESPECIFICA FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "'")
	  if temp.eof=false then
	  %>
	  <td width="27%"><font face="Verdana" size="1"><%=temp("FUNE_TX_OBS_ESPECIFICA")%></font></td>
    </tr>
    <%
	end if
        conta=conta+1
			RS.MOVENEXT        
        	LOOP
        %>
  </table>
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

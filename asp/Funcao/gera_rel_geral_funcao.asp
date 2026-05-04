<%@LANGUAGE="VBSCRIPT"%>  
<%
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
end if

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      if str_Desuso = 1 then
         str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	  else
     	 str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '2' "
	  end if	 
	end if        	  
end if

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")
str_Modulo=request("selSubModulo")
str_CdAreaAbrangencia = request("selAreaAbrangencia")
str_CdFuncao = request("selFuncao")
selG = request("selG")

str_Critica = request("chkCritica")

compl1 = ""
if str_modulo<>"0" then
	compl1=" AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & str_modulo 
end if

if str_CdFuncao <> "0" then
	compl1=compl1 & " AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_CdFuncao  & "'"
end if

str_Sub_Titulo = ""

if selG = "1" then
	compl1=compl1 + " AND FUNCAO_NEGOCIO.FUNE_TX_TP_FUN_NEG ='G'"
	str_Sub_Titulo =  str_Sub_Titulo & " - Genérica"
else
   selG = 0	
end if

if str_Critica=1 then
	compl1=compl1 + " AND FUNCAO_NEGOCIO.FUNE_TX_INDICA_CRITICA ='1'"
	str_Sub_Titulo = str_Sub_Titulo & " - Crítica"	
else
   str_Critica = 0	
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso)

if str_MegaProcesso=0 then
	IF selG=1 then
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_TX_TP_FUN_NEG ='G' " & str_usoDesuso  
	else
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO  where MEPR_CD_MEGA_PROCESSO > 0 " & str_usoDesuso 
	end if
	ssql=ssql+" ORDER BY FUNE_CD_FUNCAO_NEGOCIO "
else
	if str_modulo<>"0" then
		ssql=""
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TIPO_CLASS "
		ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
		ssql=ssql+"LEFT JOIN FUNCAO_NEGOCIO ON "
		ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
		ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
		ssql=ssql+" ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
	else
		ssql=""
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TIPO_CLASS "
		ssql=ssql+"FROM FUNCAO_NEGOCIO "
		ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
		ssql=ssql+" ORDER BY FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	end if
	
	'ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
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
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="gera_rel_geral_funcao.asp?selG=<%=selG%>&chkCritica=<%=str_Critica%>&selMegaProcesso=<%=str_MegaProcesso%>&amp;selSubModulo=<%=str_Modulo%>&amp;opt=1&selG=<%=selG%>&chkEmUso=<%=str_Uso%>&chkEmDesuso=<%=str_Desuso%>&selAreaAbrangencia=<%=str_CdAreaAbrangencia%>&selFuncao=<%=str_CdFuncao%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></b></font></td>
          <td width="27"></td>
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
    Relatório Geral de Fun&ccedil;&atilde;o R/3 <%=str_Sub_Titulo%></b></font></p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 (<font face="Verdana">Clique no código da Fun&ccedil;&atilde;o R/3 para exibir
 seus dados)</font></font></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <%
        conta=0
        DO UNTIL RS.EOF=TRUE
		   str_Imprimir = 1
		   if str_CdAreaAbrangencia <> "0" then
		      str_SQL = ""
		      str_SQL = str_SQL & " SELECT dbo.ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO "
              str_SQL = str_SQL & " FROM  dbo.FUN_NEG_ORG_AGLU INNER JOIN "
              str_SQL = str_SQL & " dbo.ORGAO_AGLUTINADOR ON dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = dbo.ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "
              str_SQL = str_SQL & " WHERE (dbo.FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO = '" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "')"
		      str_SQL = str_SQL & " and  dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = " & str_CdAreaAbrangencia
		      set rdsAreaAbran = db.Execute(str_SQL)
		      if rdsAreaAbran.EOF then
			     str_Imprimir = 0
			  end if	 
           end if
		   if str_Imprimir = 1 then
		%>
        
  <table border="0" width="74%" height="76">
    <tr> 
      <td width="12%" height="19"> </td>
      <td width="16%" height="19" bgcolor="#E0E0E0"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099"><b>Código</b></font></td>
      <td width="90%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="1"><a href="exibe_dados_funcao.asp?selMegaProcesso=0&selFuncao=<%=RS("FUNE_CD_FUNCAO_NEGOCIO")%>"><b><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%></b></a></font></td>
    </tr>
    <tr> 
      <td width="12%" height="19"></td>
      <td width="16%" height="19" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Título</b></font></td>
      <td width="90%" height="19"><font face="Verdana" color="#330099" size="1"><%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <tr> 
      <td width="12%" height="6" valign="top"></td>
      <td width="16%" height="6" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Descrição</b></font></td>
      <td width="90%" height="6" valign="top"><font face="Verdana" color="#330099" size="1"><%=RS("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <tr> 
      <td height="6" valign="top"></td>
      <td height="6" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Assuntos</b></font></td>
	  <%
	  	assuntos=""
	  
		ssql=""
		ssql="SELECT DISTINCT SUB_MODULO.SUMO_TX_DESC_SUB_MODULO FROM FUNCAO_NEGOCIO_SUB_MODULO"	  
		ssql=ssql+" INNER JOIN SUB_MODULO ON FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = SUB_MODULO.SUMO_NR_CD_SEQUENCIA"
		ssql=ssql+" WHERE FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "' "	  
		ssql=ssql+" ORDER BY SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"	  
		
		set rsassunto=db.execute(ssql)
		
		do until rsassunto.eof=true
			assuntos=assuntos & rsassunto("SUMO_TX_DESC_SUB_MODULO") & ", "		
			rsassunto.movenext
		loop
		
		assuntos=left(assuntos,len(assuntos)-2)
		
	  %>
      <td height="6" valign="top"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=assuntos%></font></td>
    </tr>
	
	<td height="6" valign="top"></td>
      <td height="6" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>&Aacute;rea 
        de Abrang&ecirc;ncia</b></font></td>
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
      <td height="6" valign="top"><font color="#330099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_AreaAbrang%> </font></td>
    </tr>
	
	<%	
	if trim(rs("FUNE_TX_TIPO_CLASS")) = "0" then
		str_TipoClass = ""
	elseif trim(rs("FUNE_TX_TIPO_CLASS")) = "1" then
		str_TipoClass = "EBP"
	elseif trim(rs("FUNE_TX_TIPO_CLASS")) = "2" then
		str_TipoClass = "EMERGE"
	end if
	
	'if str_TipoClass <> "" then
	%>	
    <tr>
      <td height="6" valign="top"></td>
      <td height="6" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Tipo de Classificação</b></font></td>
      <td height="6" valign="top"><font color="#330099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_TipoClass%></font></td>
    </tr>
    <%'end if%> 
      
  </table>
 <P>
        <%
        conta=conta+1
		end if 'if str_Imprimir = 1 then
			RS.MOVENEXT        
        	LOOP
        %>
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
<%@LANGUAGE="VBSCRIPT"%> 
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Opt = request("txtOpt")
'response.Write("<P> =====" & str_Opt & "////////////<P>" )
if request("txtMegaProcesso") <> "" then
   str_MegaProcesso = request("txtMegaProcesso")
else
   str_MegaProcesso = "0"
   if request("selMegaProcesso") <> "" then
      str_MegaProcesso = request("selMegaProcesso")      
   end if	  
end if   

if request("txtSubModulo") <> "" then
   str_Modulo = request("txtSubModulo")
else
   str_Modulo = "0"
   if request("selSubModulo") <> "" then
      str_Modulo = request("selSubModulo")      
   end if	  
end if   

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
str_Funcao=request("selFuncao")

'response.Write("<p>" & str_MegaProcesso)
'response.Write("<p>" & str_Modulo)
'response.Write("<p>" & str_Uso)
'response.Write("<p>" & str_Desuso)
'response.Write("<p>" & str_Funcao & "<p>")

'========================= SELECIONA O MEGA ================================
str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " MEPR_TX_DESCRICAO "
str_SQL = str_SQL & " ,MEPR_TX_INDICA_SUB_MODULO "
str_SQL = str_SQL & " ,MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL = str_SQL & " ,MEPR_CD_MEGA_PROCESSO "
str_SQL = str_SQL & " FROM MEGA_PROCESSO "
str_SQL = str_SQL & " WHERE  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
'response.Write(str_SQL)

set rdsMegaProcesso = db.Execute(str_SQL)
if not rdsMegaProcesso.EOF then
   str_DescMegaProcesso = rdsMegaProcesso("MEPR_TX_DESCRICAO")
   str_DsMegaProcesso = rdsMegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
   str_Possui_SubModulo = rdsMegaProcesso("MEPR_TX_INDICA_SUB_MODULO")
else
   str_DescMegaProcesso = ""
   str_DsMegaProcesso = ""
   str_Possui_SubModulo = "0"
end if
rdsMegaProcesso.close
set rdsMegaProcesso = Nothing
'========================================= SELECIONA ORIEMTAÇÃO DO MEGA ==================================
str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " MEPR_CD_MEGA_PROCESSO "
str_SQL = str_SQL & " , ORIE_NR_SEQUENCIAL"
str_SQL = str_SQL & " , ORIE_TX_ORIENTACOES"
str_SQL = str_SQL & " , ORIE_NR_ORDENACAO"
str_SQL = str_SQL & " FROM FUN_ORIEN_MEGA"
str_SQL = str_SQL & " WHERE  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL = str_SQL & " order by ORIE_NR_ORDENACAO "
'response.Write(str_SQL)
set rdsOrientMega = db.Execute(str_SQL)
'=================================================== TERMO DO MEGA ======================================
str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " MEPR_CD_MEGA_PROCESSO"
str_SQL = str_SQL & " , ORTE_NR_SEQUENCIAL"
str_SQL = str_SQL & " , ORTE_TX_TERMO"
str_SQL = str_SQL & " , ORTE_TX_DESCRICAO"
str_SQL = str_SQL & " FROM  FUN_ORIEN_MEGA_TERMOS"
str_SQL = str_SQL & " WHERE  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
'response.Write(str_SQL)
set rdsOrientMegaTermo = db.Execute(str_SQL)
'================================================== SUB MÓDULO =================================================
str_SQL_SubModulo =""
str_SQL_SubModulo = str_SQL_SubModulo & " SELECT SUMO_NR_CD_SEQUENCIA"
str_SQL_SubModulo = str_SQL_SubModulo & " ,SUMO_TX_DESC_SUB_MODULO"
str_SQL_SubModulo = str_SQL_SubModulo & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
str_SQL_SubModulo = str_SQL_SubModulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_SQL_SubModulo = str_SQL_SubModulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
str_SQL_SubModulo = str_SQL_SubModulo + " ORDER BY SUMO_TX_DESC_SUB_MODULO"
set rdsModulos = db.Execute(str_SQL_SubModulo)
'================================================== SUB MÓDULO DO MEGA =========================================
str_Sql_SubModulo = ""
str_Sql_SubModulo = str_Sql_SubModulo & " Select "
str_Sql_SubModulo = str_Sql_SubModulo & " FUN_ORIEN_MEGA_MODULO.MEPR_CD_MEGA_PROCESSO "
str_Sql_SubModulo = str_Sql_SubModulo & " , FUN_ORIEN_MEGA_MODULO.SUMO_NR_CD_SEQUENCIA "
str_Sql_SubModulo = str_Sql_SubModulo & " , FUN_ORIEN_MEGA_MODULO.ORTE_TX_DESCRICAO "
str_Sql_SubModulo = str_Sql_SubModulo & " , SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
str_Sql_SubModulo = str_Sql_SubModulo & " FROM dbo.FUN_ORIEN_MEGA_MODULO INNER JOIN dbo.SUB_MODULO ON dbo.FUN_ORIEN_MEGA_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
str_Sql_SubModulo = str_Sql_SubModulo & " where  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_Sql_SubModulo = str_Sql_SubModulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.Write(str_SQL)
set rdsOrientMegaModulo = db.Execute(str_Sql_SubModulo)
' ================================================================ FUNÇÕES ================================================
if str_MegaProcesso <> "0" then
   str_Sql_MegaProcesso2 = " and FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if
if str_modulo <> "0" then
	str_Sql_SubModulo2 = " AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & str_modulo 
end if

if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   

if str_Uso = 1 and str_Desuso = 1 then
   str_Sql_usoDesuso2 =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_Sql_usoDesuso2 =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_Sql_usoDesuso2 =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

if str_Funcao <> "0" and str_Funcao <> ""  then
	str_Sql_Funcao2 =" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao  & "'"
end if

'response.Write("<p>" & str_Sql_MegaProcesso2)
'response.Write("<p>" & str_Sql_SubModulo2)
'response.Write("<p>" & str_Sql_usoDesuso2)
'response.Write("<p>" & str_Sql_Funcao2 & "<p>" )

if str_Opt <> "RM" then
   str_Sql_SubModulo2 = ""
   str_Sql_usoDesuso2 = ""
   str_Sql_Funcao2 = ""
end if

str_Sql = ""
str_Sql = str_Sql & " Select "
str_Sql = str_Sql & " FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO "
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO.FUNE_TX_OBS_ESPECIFICA"
str_Sql = str_Sql & " ,FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO"
str_Sql = str_Sql & " FROM FUNCAO_NEGOCIO"
str_Sql = str_Sql & " INNER JOIN FUNCAO_NEGOCIO_SUB_MODULO ON"
str_Sql = str_Sql & " FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " where FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO > 0 " & str_Sql_MegaProcesso2 & str_Sql_SubModulo2 & str_Sql_usoDesuso2 & str_Sql_Funcao2 
str_Sql = str_Sql & " order by FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO, FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "

'response.Write(str_Sql)
'set rdsFuncao = db.Execute(str_SQL)
set RS = db.Execute(str_Sql)

'============================================== AREA DE ABRANGENCIA ===========================================
'str_SQL_Area = ""
'str_SQL_Area = str_SQL_Area & " Select "
'str_SQL_Area = str_SQL_Area & " dbo.FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO "
'str_SQL_Area = str_SQL_Area & " , dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO "
'str_SQL_Area = str_SQL_Area & " ,dbo.ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO"
'str_SQL_Area = str_SQL_Area & " FROM  dbo.FUN_NEG_ORG_AGLU INNER JOIN dbo.ORGAO_AGLUTINADOR ON dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = dbo.ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "
'str_SQL_Area = str_SQL_Area & " where dbo.FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO = " & str_FuncNegocio 
'set rdsFuncArea = db.Execute(str_SQL)

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
      <td colspan="3" height="20">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="90%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      
    <td width="90%" height="40" bgcolor="#333333"> 
      <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Orienta&ccedil;&otilde;es 
        ao Mapeamento de Usu&aacute;rio</strong></font></div></td>
      <td width="5%">&nbsp;</td>
    </tr>
  </table>
  
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td width="89%"></td>
  </tr>
  <tr> 
    <td bgcolor="#666666"> <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=str_DsMegaProcesso%></strong></font></div></td>
  </tr>
  <tr> 
    <td bgcolor="#999999"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Descri&ccedil;&atilde;o 
        do Mega-Processo</strong></font></div></td>
  </tr>
  <tr> 
    <td><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_DescMegaProcesso%></font></div></td>
  </tr>
  <tr> 
    <td bgcolor="#999999"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Orienta&ccedil;&otilde;es 
        Gerais para o Mega-Processo</strong></font></div></td>
  </tr>
  <% do while not rdsOrientMega.EOF %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrientMega("ORIE_NR_SEQUENCIAL")%> 
      - <%=rdsOrientMega("ORIE_TX_ORIENTACOES")%></font></td>
  </tr>
  <% rdsOrientMega.movenext
  Loop %>
</table> 
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#999999"> 
    <td colspan="2"> 
      <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Termos 
        Novos Relevantes</strong></font></div></td>
  </tr>
  <tr> 
    <td width="25%"><div align="center"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Termo</strong></font></div></td>
    <td width="65%"><div align="left"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Descri&ccedil;&atilde;o</strong></font></div></td>
  </tr>
  <% do while not rdsOrientMegaTermo.EOF %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrientMegaTermo("ORTE_TX_TERMO")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrientMegaTermo("ORTE_TX_DESCRICAO")%></font></td>
  </tr>
  <% rdsOrientMegaTermo.movenext
  Loop %>
</table>
<% 'if str_Possui_SubModulo = "1" and not rdsOrientMegaModulo.EOF then %>
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#999999"> 
    <td colspan="2"> <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Assuntos 
        Relacionados ao Mega</strong></font></div></td>
  </tr>
  <tr> 
    <td width="25%"><div align="center"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Assunto</strong></font></div></td>
    <td width="65%"><div align="left"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Descri&ccedil;&atilde;o</strong></font></div></td>
  </tr>
  <% do while not rdsOrientMegaModulo.EOF %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrientMegaModulo("SUMO_TX_DESC_SUB_MODULO")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrientMegaModulo("ORTE_TX_DESCRICAO")%></font></td>
  </tr>
  <% rdsOrientMegaModulo.movenext
  Loop %>
</table>
<% 'end if %>
<table width="90%" height="129" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr valign="top" bgcolor="#999999"> 
    <td height="19" colspan="2"> <p align="center" style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Orienta&ccedil;&otilde;es 
        por Fun&ccedil;&atilde;o</strong></font></td>
  </tr>
  <%
  conta=0
  str_SubModulo_Anterior = 99999
  DO UNTIL RS.EOF=TRUE
     str_Contador = str_Contador + 1
	'If not isNull(RS("SUMO_NR_CD_SEQUENCIA")) and str_SubModulo_Anterior <> RS("SUMO_NR_CD_SEQUENCIA")  then
     str_N = "  entrou 1 "
	 If str_SubModulo_Anterior <> RS("SUMO_NR_CD_SEQUENCIA")  then
       str_N = "  entrou 2 "
       str_SubModulo_Anterior = RS("SUMO_NR_CD_SEQUENCIA") 
	      str_N = "  entrou 3 "
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
		  
	 	  str_SQL_MegaProc = ""
          str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
          str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
          str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
          str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_INDICA_SUB_MODULO "		
          str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
          str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO = " & RS("MEPR_CD_MEGA_PROCESSO")
          str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
		  set rdsMegaProcesso = db.Execute(str_SQL_MegaProc)
		  if not rdsMegaProcesso.EOF then
		     str_DsMegaProcesso = rdsMegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
		     str_Indica_SubModulo = rdsMegaProcesso("MEPR_TX_INDICA_SUB_MODULO")		  		   
		  else
		     str_DsMegaProcesso = "NÃO ENCONTRADO"
		     str_Indica_SubModulo = ""
		  end if
		  rdsMegaProcesso.close
		  set rdsMegaProcesso = Nothing
		  if str_Indica_SubModulo = "1" then
		     str_Texto_Sub_Ass = "Assunto :"
		  elseif str_Indica_SubModulo = "0" then
		     str_Texto_Sub_Ass = "Assunto :"
		  ELSE
		     str_Texto_Sub_Ass = ""   		
		  end if
	%>
    <tr valign="top" bgcolor="#999999"> 
    <td height="19" colspan="2"> <p align="center" style="margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2" face="Verdana"><strong><font size="3"><%=str_Texto_Sub_Ass%></font></strong></font> <font color="#FFFFFF" size="2" face="Verdana"><strong><font size="3"><%=str_DsSubModulo%></font></strong></font></td>
  </tr>
  <% end if %>
  <tr valign="top"> 
    <td width="17%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">Mega</font></td>
    <td width="83%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="1"><%=str_DsMegaProcesso%> <%'=RS("SUMO_NR_CD_SEQUENCIA")%>
<%'=str_N%></font></strong></td>
  </tr>
  <tr valign="top"> 
    <td height="19"><font face="Verdana" size="2">Fun&ccedil;&atilde;o</font></td>
    <td height="19"><strong><font face="Verdana" size="1"><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%> </font></strong></td>
  </tr>
  <tr valign="top"> 
    <td width="17%" height="19"><font face="Verdana" size="1">&nbsp;<font face="Verdana" size="2">Descrição</font></font></td>
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
    <td height="19"><font face="Verdana" size="2">&Aacute;rea de Abrang&ecirc;ncia</font></td>
    <td height="19"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_AreaAbrang%></font></td>
  </tr>
  <tr valign="top"> 
    <td height="19"><font face="Verdana" size="2">Observa&ccedil;&atilde;o Espec&iacute;fica 
      para Fun&ccedil;&atilde;o</font></td>
    <td height="19"><font face="Verdana" size="1"><%=RS("FUNE_TX_OBS_ESPECIFICA")%></font></td>
  </tr>
  <tr valign="top" bgcolor="#0000FF"> 
    <td height="5" bgcolor="#CCCCCC"></td>
    <td height="5" bgcolor="#CCCCCC"></td>
  </tr>
  <%
        conta=conta+1
			RS.MOVENEXT       
        	LOOP
        %>
</table>
</body>
</html>

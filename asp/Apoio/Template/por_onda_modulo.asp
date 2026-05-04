<!--#include file="../conn_consulta.asp" -->
<html>
<%
server.ScriptTimeout=99999999

opti=request("op")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

onda = request("onda")
modulo = request("modulo")

ssql=""
ssql="SELECT DISTINCT dbo.APOIO_LOCAL_ONDA.ONDA_CD_ONDA, dbo.ONDA.ONDA_TX_DESC_ONDA FROM dbo.APOIO_LOCAL_ONDA "
ssql=ssql+"INNER JOIN dbo.ONDA ON dbo.APOIO_LOCAL_ONDA.ONDA_CD_ONDA = dbo.ONDA.ONDA_CD_ONDA WHERE (dbo.APOIO_LOCAL_ONDA.APLO_NR_ATRIBUICAO=1) "
ssql=ssql+"ORDER BY dbo.ONDA.ONDA_TX_DESC_ONDA"

set fonte=db.execute(ssql)

%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="88%" border="0">
  <tr> 
    <td width="68%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
        POR ONDA X ASSUNTO - APOIADOR LOCAL</strong></font></div></td>
    <td width="32%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<table width="754" border="0">
<%
do until fonte.eof=true

if trim(onda)=trim(fonte("ONDA_CD_ONDA")) then
	seleciona=1
	figura="baixo"
else
	seleciona=0
	figura="lado"
end if
%>
  <tr bgcolor="#FFFFFF"> 
    <td width="26"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte("ONDA_CD_ONDA")%>"><img src="<%=figura%>.jpg" width="17" height="20"></a></font></td>
    <td width="7"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td height="28" colspan="4" width="512"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_onda_modulo.asp?op=<%=opti%>&onda=<%=fonte("ONDA_CD_ONDA")%>#<%=fonte("ONDA_CD_ONDA")%>"><%=fonte("ONDA_TX_DESC_ONDA")%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="66"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
<%
if figura="baixo" then

ssql=""
ssql="SELECT DISTINCT "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA0, "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 8),2)) AS FAIXA1 , "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 5),2)) AS FAIXA2 , "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA3, "
ssql=ssql+"dbo.MEGA_PROCESSO.MEPR_TX_ABREVIA AS FAIXA4 , dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, "
ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA,  dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM dbo.APOIO_LOCAL_MODULO INNER JOIN dbo.SUB_MODULO ON "
ssql=ssql+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"RIGHT JOIN APOIO_LOCAL_MULT ON "
ssql=ssql+"dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
ssql=ssql+"INNER JOIN APOIO_LOCAL_ONDA ON "
ssql=ssql+"dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ONDA.USMA_CD_USUARIO "
ssql=ssql+"LEFT JOIN dbo.MEGA_PROCESSO ON "
ssql=ssql+"(CONVERT(char(3),dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO)) IN ((REPLACE((REPLACE(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS,'0','')),'-',', '))) "
ssql=ssql+"WHERE APOIO_LOCAL_ONDA.ONDA_CD_ONDA= " & onda & " AND APOIO_LOCAL_ONDA.APLO_NR_ATRIBUICAO = 1 "
ssql=ssql+"ORDER BY FAIXA0, FAIXA1, FAIXA2, FAIXA3, FAIXA4, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

set fonte2=db.execute(ssql)

do until fonte2.eof=true

if trim(modulo)=trim(fonte2("SUMO_NR_CD_SEQUENCIA")) then
	seleciona2=1
	figura2="baixo"
else
	seleciona2=0
	figura2="lado"
end if
	
%>  
  <tr bgcolor="#FFFFFF"> 
    <td width="26"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="7"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
	<%
	IF NOT ISNULL(fonte2("FAIXA4")) THEN
	%>
	<td height="28" colspan="4" width="512"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><img src="<%=figura2%>.jpg" width="17" height="20"></a></font><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_onda_modulo.asp?op=<%=opti%>&modulo=<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>&onda=<%=fonte("ONDA_CD_ONDA")%>#<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><%=fonte2("FAIXA4")%> - <%=fonte2("SUMO_TX_DESC_SUB_MODULO")%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
	<%END IF%>
    <%
	IF (fonte2("FAIXA0")=fonte2("FAIXA1")) AND (fonte2("FAIXA0")=fonte2("FAIXA2")) AND (fonte2("FAIXA0")<>fonte2("FAIXA3")) THEN
	%>
	<td height="28" colspan="4" width="512"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><img src="<%=figura2%>.jpg" width="17" height="20"></a></font><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_onda_modulo.asp?op=<%=opti%>&modulo=<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>&onda=<%=fonte("ONDA_CD_ONDA")%>#<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><%=fonte2("FAIXA0")%>-<%=fonte2("FAIXA3")%> - <%=fonte2("SUMO_TX_DESC_SUB_MODULO")%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
	<%END IF%>
    <%
	IF (fonte2("FAIXA0")=fonte2("FAIXA1")) AND (fonte2("FAIXA0")<>fonte2("FAIXA2")) AND (fonte2("FAIXA0")<>fonte2("FAIXA3")) THEN
	%>
	<td height="28" colspan="4" width="512"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><img src="<%=figura2%>.jpg" width="17" height="20"></a></font><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_onda_modulo.asp?op=<%=opti%>&modulo=<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>&onda=<%=fonte("ONDA_CD_ONDA")%>#<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><%=fonte2("FAIXA0")%>-<%=fonte2("FAIXA2")%>-<%=fonte2("FAIXA3")%> - <%=fonte2("SUMO_TX_DESC_SUB_MODULO")%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
	<%END IF%>
	<%
	IF (fonte2("FAIXA0")<>fonte2("FAIXA1")) AND (fonte2("FAIXA0")<>fonte2("FAIXA2")) AND (fonte2("FAIXA0")<>fonte2("FAIXA3")) THEN
	%>
	<td height="28" colspan="4" width="512"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><img src="<%=figura2%>.jpg" width="17" height="20"></a></font><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_onda_modulo.asp?op=<%=opti%>&modulo=<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>&onda=<%=fonte("ONDA_CD_ONDA")%>#<%=fonte2("SUMO_NR_CD_SEQUENCIA")%>"><%=fonte2("FAIXA0")%>-<%=fonte2("FAIXA1")%>-<%=fonte2("FAIXA2")%>-<%=fonte2("FAIXA3")%> - <%=fonte2("SUMO_TX_DESC_SUB_MODULO")%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
	<%END IF%>
	<td width="66"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
<%
if seleciona2=1 then
	
	ssql=""
	ssql="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR AS LOTACAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO, dbo.APOIO_LOCAL_ONDA.ONDA_CD_ONDA "
	ssql=ssql+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "	
	ssql=ssql+"dbo.APOIO_LOCAL_ONDA ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ONDA.USMA_CD_USUARIO "
	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_ONDA.ONDA_CD_ONDA = " & onda & ") AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & modulo & ") "
	ssql=ssql+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"
	
	'response.write ssql

	set rs=db.execute(ssql)
	
  %>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2" bgcolor="#FFFFFF" width="39"><font color="#FFFFFF">&nbsp;</font></td>
    <td width="230" height="44"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></td>
    <td width="71" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="117" bgcolor="#CCCCCC">
      <p align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Lotação</font></strong></td>
    <td width="76" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="66" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
    <td width="258" bgcolor="#CCCCCC"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Assunto</font></strong></div></td>	
    <td width="111" bgcolor="#CCCCCC"><div align="left"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">&Oacute;rgao 
        Apoiado</font></strong></div></td>
  </tr>
  <%
  do until rs.eof=true
  if cor="white" then
  	cor="#DCDCED"
  else
  	cor="white"
  end if
  %>
  <tr valign="middle" bgcolor="<%=cor%>"> 
    <td colspan="2" bgcolor="#FFFFFF" width="39"><font color="#FFFFFF">&nbsp;</font></td>
    <td height="26" width="230" valign="top"> <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre></td>

    <td width="71" valign="top"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre>
      </div></td>
      <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
    if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
    %>
    <td width="117" valign="top"> 
      <%ON ERROR RESUME NEXT %>
	  <p align="left"><font size="1" face="Arial, Helvetica, sans-serif"><%=temp("lotacao")%></font></td>
    <td width="76" valign="top"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
    <td width="66" valign="top"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("RAMAL")%></font></pre>
      </div></td>
    <%
		ssql=""
		ssql="SELECT distinct dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS NOME "
		ssql=ssql+"FROM dbo.SUB_MODULO INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON "
		ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
		ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = '" & rs("chave") & "') AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO=1) "
		ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"
		
		set rs_modulo=db.execute(ssql)
		
		modulo=""
		do until rs_modulo.eof=true
			modulo=modulo & "," & rs_modulo("NOME")		
			rs_modulo.movenext
		loop
		
		if len(modulo)>1 then
			modulo=right(modulo,len(modulo)-1)
		end if
    %>
    <td valign="top" width="258"> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=modulo%></font></div></td>
    <%
    	ssql=""
    	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR AS APOIO FROM APOIO_LOCAL_ORGAO WHERE (USMA_CD_USUARIO ='" & UCASE(rs("chave")) & "') AND (APLO_NR_ATRIBUICAO=1) AND SUBSTRING(ORME_CD_ORG_MENOR,11,5) = '00000' ORDER BY ORME_CD_ORG_MENOR"
    	
		set rs_orgao=db.execute(ssql)    
    
    	orgao=""
    	do until rs_orgao.eof=true
    		set temp2=db.execute("SELECT ORME_SG_ORG_MENOR AS APOIADO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs_orgao("APOIO") & "'")
			if temp2.eof=true then
			    set temp2=db.execute("SELECT AGLU_SG_AGLUTINADO AS APOIADO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs_orgao("APOIO") & "'")
			end if
			orgao=orgao  & "," & trim(temp2("APOIADO"))
    		rs_orgao.movenext
    	loop
    	
    	if len(orgao)>1 then
			orgao=right(orgao,len(orgao)-1)
		end if
    %>
    <td valign="top" width="111"> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=orgao%></A></div></td>
  </tr>
  <%
  tem=tem+1
  rs.movenext
  loop
  end if
  if tem<>0 then
  %>
  <tr valign="middle" bgcolor="#FFFFFF"> 
    <td colspan="2" width="39">&nbsp;</td>
    <td height="26" align="center" valign="middle" bgcolor="#CCCCCC" width="230"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
      de Registros : </strong><font size="2"><%=tem%></font></font></td>
    <td width="71">&nbsp;</td>
    <td width="117"></td>
    <td width="76">&nbsp;</td>
    <td width="66">&nbsp;</td>
    <td valign="top" width="111">&nbsp;</td>
  </tr>
  <%
  end if
  tem=0
  fonte2.movenext
  loop
  end if
  seleciona=0
  fonte.movenext
  loop
  %>
</table>
<br><br>
<p>&nbsp;</p>
</body>
</html>


<!--#include file="../conn_consulta.asp" -->
<%
if request("excel")=1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<html>
<%
server.ScriptTimeout=99999999

tem=0
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation=3

modulo=request("selModulo_")

if len(modulo)<1 then
	modulo=0
end if

aglu=request("str01")
orgao2=request("str02")
orgao3=request("str03")

if aglu>0 and orgao2=0 and orgao3=0 then
	set rs=db.execute("SELECT AGLU_SG_AGLUTINADO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & aglu)
	txtaglu=rs("AGLU_SG_AGLUTINADO")
	set fonte = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO WHERE ORME_CD_ORG_MENOR LIKE '" & aglu & "%' AND APLO_NR_ATRIBUICAO=" & request("Atrib") & " ORDER BY ORME_CD_ORG_MENOR")
end if

if aglu>0 and orgao2>0 and orgao3=0 then
	aglu=orgao2
	set rs=db.execute("SELECT ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & aglu & "00000000'")
	txtaglu=rs("ORME_SG_ORG_MENOR")
	set fonte = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO WHERE ORME_CD_ORG_MENOR LIKE '" & aglu & "%' AND APLO_NR_ATRIBUICAO=" & request("Atrib") & " ORDER BY ORME_CD_ORG_MENOR")
end if

if aglu>0 and orgao2>0 and orgao3>0 then
	aglu=orgao3
	set rs=db.execute("SELECT ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & aglu & "00000'")
	txtaglu=rs("ORME_SG_ORG_MENOR")
	set fonte = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO WHERE ORME_CD_ORG_MENOR LIKE '" & aglu & "%' AND APLO_NR_ATRIBUICAO=" & request("Atrib") & " ORDER BY ORME_CD_ORG_MENOR")
end if

%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="67%" height="26" border="0">
           <tr>
                      <td width="5%">
                         <div align="right">
                                   <a href="javascript:history.go(-1)"><img src="../../../imagens/seta_esquerda_01.jpg" width="21" height="18" border="0"></a></div>
                      </td>
                      <td width="5%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
                      <td width="8%">
                         <div align="right">
                                   <a href="javascript:print()"><img src="../impressão.jpg" width="27" height="21" border="0"></a></div>
                      </td>
                      <td width="24%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong></td>
           </tr>
</table>
<p><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Órgão Selecionado : <%=txtaglu%></strong></font></p>
<%

select case request("visual")
case 1
%> <p></p>
<%
ssql=""
ssql="SELECT DISTINCT Left([ORME_CD_ORG_MENOR],2) AS ORGAO, APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, APOIO_LOCAL_ORGAO.USMA_CD_USUARIO"
ssql=ssql+" FROM APOIO_LOCAL_ORGAO "
ssql=ssql+"WHERE (((Left([ORME_CD_ORG_MENOR],2))=" & aglu & ") AND ((APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO)=" & request("Atrib") & "))"

set quantos=db.execute(ssql)

qtos=quantos.recordcount

%> <table width="85%" border="0">
           <tr bgcolor="#cccccc">
                      <td width="28%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ÓRGÃO APOIADO</font></strong></td>
                      <td colspan="2"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ASSUNTO</font></strong><font size="2"> </font>
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <%
do until fonte.eof=true
%> <tr>
                      <%
	if len(fonte("ORME_CD_ORG_MENOR"))=2 then
		set orgao=db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & fonte("ORME_CD_ORG_MENOR"))
		pre=" - SEDE"
	else
		set orgao=db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("ORME_CD_ORG_MENOR") & "'")	
		pre=""
	end if	
	
org_=orgao("ORGAO")

ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      COUNT(DISTINCT dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO) AS CONTA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND "

if modulo=0 then
	ssql=ssql+"                     (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"
else
	ssql=ssql+"                     (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"
end if

set rs_org=db.execute(ssql)

CONTA=0

if rs_org.eof=false then
	%> <td bgcolor="#d1d1d1"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=org_%></font></strong></td>
                      <td width="66%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td width="6%">
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <%
end if

do until rs_org.eof=true
%> <tr>
                      <td><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td bgcolor="#e6e6e6"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("SUMO_TX_DESC_SUB_MODULO")%></font></strong></td>
                      <td>
                         <div align="center">
                                   <font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("CONTA")%></font></div>
                      </td>
           </tr>
           <%
  CONTA=CONTA+rs_org("CONTA")
  tem=tem+1
rs_org.movenext
loop
if tem>0 then
%> <tr>
                      <td><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      <td>
                         <div align="right">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total por Órgão Apoiado : </strong></font>
                         </div>
                      </td>
                      <td>
                         <div align="center">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=CONTA%></b></font></div>
                      </td>
           </tr>
           <%
 end if
fonte.movenext
loop
%> </table>
<%
case 2
%> <table width="93%" border="0">
           <tr bgcolor="#cccccc">
                      <td width="23%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ÓRGÃO APOIADO</font></strong></td>
                      <td colspan="5"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ASSUNTO</font></strong><font size="2"> </font>
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <%
do until fonte.eof=true
%> <tr>
                      <%
	if len(fonte("ORME_CD_ORG_MENOR"))=2 then
		set orgao=db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & fonte("ORME_CD_ORG_MENOR"))
	else
		set orgao=db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("ORME_CD_ORG_MENOR") & "'")	
	end if	
	
org_=orgao("ORGAO")

ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      COUNT(DISTINCT dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO) AS CONTA, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND "

if modulo=0 then
	ssql=ssql+"                     (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"
else
	ssql=ssql+"                     (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"
end if

set rs_org=db.execute(ssql)

CONTA=0

if rs_org.eof=false then
	%> <td bgcolor="#d1d1d1"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=org_%></font></strong></td>
                      <td colspan="4"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td width="8%">
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <%
end if

do until rs_org.eof=true
%> <tr>
                      <td><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td colspan="4" bgcolor="#e6e6e6"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("SUMO_TX_DESC_SUB_MODULO")%></font></strong></td>
                      <td>
                         <div align="center">
                         </div>
                      </td>
                      <%
	CONTADOR = rs_org("CONTA")
	CONTA=CONTA+rs_org("CONTA")
	%> </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td width="8%" bgcolor="#ffffff">&nbsp;</td>
                      <td width="32%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></td>
                      <td width="21%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lotação</strong></font></td>
                      <td width="8%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Chave</strong></font></td>
                      <td>&nbsp;</td>
           </tr>
           <%
ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO AS CHAVE, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR "
ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND "
ssql=ssql+"                      (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "') AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs_org("SUMO_NR_CD_SEQUENCIA") & ") ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

set nomes=db.execute(ssql)

do until nomes.eof=true
%> <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("USMA_TX_NOME_USUARIO")%></font></td>
                      <%
	SET TEMP=DB.EXECUTE("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & nomes("LOTACAO") & "'")
	LOT=TEMP("ORGAO")
	%> <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=LOT%></font></td>
                      <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("CHAVE")%></font></td>
                      <td>&nbsp;</td>
           </tr>
           <%
    tem=tem+1
nomes.movenext
loop
%> <tr>
                      <td>&nbsp;</td>
                      <td colspan="4" bgcolor="#ffffff">
                         <div align="right">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total por Assunto: </strong></font>
                         </div>
                      </td>
                      <td bgcolor="#ffffff">
                         <div align="center">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=CONTADOR%></strong></font></div>
                      </td>
           </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <%
rs_org.movenext
loop
%> <%
  if tem>0 then
  %> <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <%end if%> <%
response.Flush()
fonte.movenext
loop
%> 

<%
case 3
%>

</table>
<table width="93%" border="0">
           <tr bgcolor="#cccccc">
                      <td width="23%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ÓRGÃO APOIADO</font></strong></td>
                      <td colspan="5"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">ASSUNTO</font></strong><font size="2"> </font>
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <tr>
                      <%
if len(aglu)=2 then
	set orgao=db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & aglu)
else
	set orgao=db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & left(aglu & "00000000000000",15) & "'")	
end if	
	
org_=orgao("ORGAO")

ssql=""
ssql="SELECT     DISTINCT dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO LIKE '%2') AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND "

if modulo=0 then
	ssql=ssql+"                     (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & aglu & "%')"
else
	ssql=ssql+"                     (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & aglu & "%')"
end if

set rs_org=db.execute(ssql)

CONTA=0

if rs_org.eof=false then
	%> <td bgcolor="#d1d1d1"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=org_%></font></strong></td>
                      <td colspan="4"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td width="8%">
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <%
end if

do until rs_org.eof=true
%> <tr>
                      <td><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                      <td colspan="4" bgcolor="#e6e6e6"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("SUMO_TX_DESC_SUB_MODULO")%></font></strong></td>
                      <td>
                         <div align="center">
                         </div>
                      </td>
           </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td width="8%" bgcolor="#ffffff">&nbsp;</td>
                      <td width="32%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></td>
                      <td width="21%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lotação</strong></font></td>
                      <td width="8%" bgcolor="#dadada"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Chave</strong></font></td>
                      <td>&nbsp;</td>
           </tr>
           <%
ssql=""
ssql="SELECT     DISTINCT dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
'ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO AS CHAVE, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO, dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR "
ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO LIKE '%2') AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("Atrib") & ") AND "
ssql=ssql+"                      (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & aglu & "%') AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs_org("SUMO_NR_CD_SEQUENCIA") & ") ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

set nomes=db.execute(ssql)

CONTADOR = nomes.recordcount
CONTA=CONTA + nomes.recordcount

do until nomes.eof=true

	%> <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("USMA_TX_NOME_USUARIO")%></font></td>
                      <%
	SET TEMP=DB.EXECUTE("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & nomes("LOTACAO") & "'")
	LOT=TEMP("ORGAO")
	%> <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=LOT%></font></td>
                      <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("CHAVE")%></font></td>
                      <td>&nbsp;</td>
           </tr>
           <%
    tem=tem+1
nomes.movenext
loop
%> <tr>
                      <td>&nbsp;</td>
                      <td colspan="4" bgcolor="#ffffff">
                         <div align="right">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total por Assunto: </strong></font>
                         </div>
                      </td>
                      <td bgcolor="#ffffff">
                         <div align="center">
                                   <font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=CONTADOR%></strong></font></div>
                      </td>
           </tr>
           <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <%
rs_org.movenext
loop
%> <%
  if tem>0 then
  %> <tr>
                      <td>&nbsp;</td>
                      <td colspan="4">&nbsp;</td>
                      <td>&nbsp;</td>
           </tr>
           <%end if%> <%
response.Flush()
%> </table>
<%
end select
%> <%if tem=0 then%> <p><font color="#800000"><strong>Nenhum Registro Encontrado para a Seleção</strong></font></p>
<%end if%>
</body>

</html>
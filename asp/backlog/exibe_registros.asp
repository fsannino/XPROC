<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

mega = request("selMega")

assunto = request("selModulo")

org1 = request("str01")
org2 = request("str02")
org3 = request("str03")

if org3<>0 then
	orgao = org3
else
	if org2<>0 then
		orgao=org2
	else
		orgao = org1
	end if
end if

if orgao=0 then
	tem_o=0
end if

titulo = request("txttitulo")

if len(orgao)>2  or tem_o=0 then

	if tem_o<>0 then
		ssql=""
		ssql="SELECT MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, ORGAO_MENOR.ORME_SG_ORG_MENOR AS ORGAO, BACKLOG.* "
		ssql=ssql+" FROM BACKLOG INNER JOIN MEGA_PROCESSO ON "
		ssql=ssql+" BACKLOG.MEPR_CD_MEGA_PROCESSO = MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO INNER JOIN ORGAO_MENOR ON "
		ssql=ssql+" BACKLOG.ORME_CD_ORG_MENOR = ORGAO_MENOR.ORME_CD_ORG_MENOR "
	else
		ssql=""
		ssql="SELECT MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, BACKLOG.* "
		ssql=ssql+" FROM BACKLOG INNER JOIN MEGA_PROCESSO ON "
		ssql=ssql+" BACKLOG.MEPR_CD_MEGA_PROCESSO = MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
	end if
	
	if mega<>0 then
		compl = compl + " BACKLOG.MEPR_CD_MEGA_PROCESSO=" & mega & " AND"
	end if

	if assunto<>0 then
		compl = compl + " BACKLOG.SUMO_NR_CD_SEQUENCIA=" & assunto & " AND"
	end if

	if orgao<>0 then
		compl = compl + " BACKLOG.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' AND"
	end if

	if len(titulo)<>0 then
		compl = compl + " BACKLOG.BALO_TX_TITULO LIKE '%" & titulo & "%' AND"
	end if
	
	if len(compl)>0 then
		ssql=ssql + "WHERE" + left(compl, len(compl)-4)
	end if
	
	ssql=ssql+" ORDER BY BACKLOG.BALO_TX_TITULO"

else

	ssql=""
	ssql="SELECT MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO AS ORGAO, BACKLOG.* "
	ssql=ssql+" FROM BACKLOG INNER JOIN MEGA_PROCESSO ON "
	ssql=ssql+" BACKLOG.MEPR_CD_MEGA_PROCESSO = MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO INNER JOIN ORGAO_AGLUTINADOR ON "
	ssql=ssql+" BACKLOG.ORME_CD_ORG_MENOR = ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "

	if mega<>0 then
		compl = compl + " BACKLOG.MEPR_CD_MEGA_PROCESSO=" & mega & " AND"
	end if

	if assunto<>0 then
		compl = compl + " BACKLOG.SUMO_NR_CD_SEQUENCIA=" & assunto & " AND"
	end if

	if orgao<>0 then
		compl = compl + " BACKLOG.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' AND"
	end if
	
	if len(titulo)<>0 then
		compl = compl + " BACKLOG.BALO_TX_TITULO LIKE '%" & titulo & "%' AND"
	end if
	
	if len(compl)>0 then
		ssql=ssql + "WHERE" + left(compl, len(compl)-4)
	end if

	ssql=ssql+" ORDER BY BACKLOG.BALO_TX_TITULO"

end if

set rs = db.execute(ssql)
%>
<html>
<head>
<title>#BACKLOG - Solicitações de Melhoria no SAP R/3#</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000099" alink="#000099" link="#000099">
<p>&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080">Consulta de Solicita&ccedil;&otilde;es 
  de Melhoria na Solu&ccedil;&atilde;o Configurada no SAP R/3</font> </p>
<p align="center">&nbsp;</p>
<table width="91%" border="0" height="26">
  <tr> 
    <td width="12%" height="27">&nbsp;</td>
    <td width="55%" height="27" bgcolor="#000099"> 
      <div align="left"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Solicita&ccedil;&otilde;es 
        Encontradas</b></font></div>
    </td>
    <td width="17%" height="27">&nbsp;</td>
  </tr>
<%
do until rs.eof=true

cod=rs("BALO_CD_COD_BACKLOG")

%>

  <tr> 
    <td width="12%" height="27">&nbsp;</td>
    <td height="27" colspan="2" bordercolor="#999999"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("BALO_TX_TITULO")%></font></td>
	<td height="27" colspan="2" bordercolor="#999999"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="valida_consulta_backlog.asp?registro=<%=cod%>">Exibir</a></font></b></td>
    <%
	if Session("Acesso")=1 then
	%>
    <td height="27" colspan="2" width="8%" bordercolor="#999999"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="carrega_backlog.asp?registro=<%=cod%>">Editar</a></font></b></td>
	<%
	end if
	%>
  </tr>
<%
tem=tem+1
rs.movenext
loop
%>
</table>
<%
if tem=0 then
%>
<p align="center"><b><font color="#990000">Nenhum Registro Encontrado para a Sele&ccedil;&atilde;o</font></b></p>
<p align="center">&nbsp;</p>
<%
end if
%>
</body>
</html>
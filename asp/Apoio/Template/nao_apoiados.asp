<%
'Response.Buffer = True
'Response.ContentType = "application/vnd.ms-excel"
%>
<!--#include file="../conn_consulta.asp" -->
<html>
<%
Server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs1=db.execute("SELECT * FROM ORGAO_AGLUTINADOR ORDER BY AGLU_CD_AGLUTINADO")

set rs2=db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR, ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_STATUS ='A' ORDER BY ORME_CD_ORG_MENOR")

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
<p><font color="#000099" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Rela&ccedil;&atilde;o 
  de &Oacute;rg&atilde;os N&atilde;o Apoiados</strong></font></p>
<table width="64%" border="0">
  <tr bgcolor="#000066"> 
    <td width="24%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></strong></td>
    <td width="76%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nome 
      de &Oacute;rg&atilde;o</font></strong></td>
  </tr>
  <%
  do until rs1.eof=true
  
  ssql=""
  ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
  ssql=ssql+"dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO "
  ssql=ssql+"FROM         dbo.APOIO_LOCAL_ORGAO INNER JOIN "
  ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
  ssql=ssql+"WHERE     (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & RS1("AGLU_CD_AGLUTINADO") & "') AND "
  ssql=ssql+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1)"
  
  set temp=db.execute(ssql)
  
  if temp.eof=true then
  %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs1("AGLU_CD_AGLUTINADO")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs1("AGLU_SG_AGLUTINADO")%> - SEDE</font></td>
  </tr>
  <%
  end if
  rs1.movenext
  loop

  do until rs2.eof=true

  ssql=""
  ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
  ssql=ssql+"dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO "
  ssql=ssql+"FROM         dbo.APOIO_LOCAL_ORGAO INNER JOIN "
  ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
  ssql=ssql+"WHERE     (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & RS2("ORME_CD_ORG_MENOR") & "') AND "
  ssql=ssql+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1)"
  
  set temp=db.execute(ssql)

  if temp.eof=true then
	if RS2("ORME_CD_ORG_MENOR")<>9999 then
  %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=RS2("ORME_CD_ORG_MENOR")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=RS2("ORME_SG_ORG_MENOR")%></font></td>
  </tr>
  <%
	tem=tem+1
	end if
  end if
  rs2.movenext
  loop
  %>
</table>
<p><strong>Org&atilde;os n&atilde;o apoiados :</strong><%=tem%></p>
<p>&nbsp;</p>
</body>
</html>

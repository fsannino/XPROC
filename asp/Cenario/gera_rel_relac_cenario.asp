<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("selExcel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

visual=request("selvis")

cenario="NO"

on error resume next

if request("ID")<>0 THEN
if err.number<>0 then
	cenario=request("ID")
end if
end if

if cenario="NO" then
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso"))
	pre_mega=TRIM(rs("MEPR_TX_ABREVIA"))
	compl1 = " WHERE (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO LIKE '" & pre_mega & "%')"
	compl2 = " WHERE (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE LIKE '" & pre_mega & "%')"
else
	compl1 = " WHERE (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO ='" & cenario & "')"
	compl2 = " WHERE (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE ='" & cenario & "')"
end if


if visual=1 then
	txt_visual="CENÁRIOS RELACIONADOS"
else
	txt_visual="RELACIONADO COM CENÁRIOS"
end if
%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
</script>
</head>

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


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <%if request("selExcel")<>1 then%>
</form>
<%end if%>
<p><font face="Verdana" color="#330099" size="3">Relatório de Relação
Cenário x Cenário - </font><b><font face="Verdana" color="#330099" size="1"><%=txt_visual%></font></b></p>
<%
if visual=1 then
	ssql="SELECT     dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE "
	ssql=ssql+"FROM         dbo.CENARIO_TRANSACAO INNER JOIN "
	ssql=ssql+                      "dbo.CENARIO ON dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO = dbo.CENARIO.CENA_CD_CENARIO "
	ssql=ssql+ compl1 & " AND (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE <> '') "
	ssql=ssql+" ORDER BY dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO"
	set rs=db.execute(ssql)
%>
<table border="0" width="850">
  <tr>
    <td width="409" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenário</b></font></td>
    <td width="427" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenários Relacionados</b></font></td>
  </tr>
<%
ATUAL=""
ANTERIOR=""

do until rs.eof=true

SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & UCASE(RS("CENA_CD_CENARIO")) & "'")
VALOR=TEMP("CENA_CD_CENARIO") & "-" & TEMP("CENA_TX_TITULO_CENARIO")

ATUAL=UCASE(RS("CENA_CD_CENARIO"))

IF TRIM(ANTERIOR)<>TRIM(ATUAL) THEN
	VALOR=VALOR
ELSE
	VALOR=" "
END IF

SET TEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & UCASE(RS("CENA_CD_CENARIO_SEGUINTE")) & "'")
VALOR2=TEMP2("CENA_CD_CENARIO") & "-" & UCASE(TEMP2("CENA_TX_TITULO_CENARIO"))

IF COR="WHITE" THEN
	COR="#E1E1E1"
ELSE
	COR="WHITE"
END IF
%>
  <tr>
    <td width="409"><font face="Verdana" size="2">
      <p style="margin-top: 0; margin-bottom: 0"><%=UCASE(VALOR)%></font></td>
    <td width="427" bgcolor="<%=COR%>"><font face="Verdana" size="2">
      <p style="margin-top: 0; margin-bottom: 0"><%=UCASE(VALOR2)%></font></td>
  </tr>
<%
tem=1
ANTERIOR=UCASE(RS("CENA_CD_CENARIO"))
rs.movenext
loop
%>
</table>
<%
else
	ssql="SELECT     dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE "
	ssql=ssql+"FROM         dbo.CENARIO_TRANSACAO INNER JOIN "
	ssql=ssql+                      "dbo.CENARIO ON dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO = dbo.CENARIO.CENA_CD_CENARIO "
	ssql=ssql+ compl2 & " AND (dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE <> '')"
	ssql=ssql+" ORDER BY dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO"
	set rs=db.execute(ssql)
%>
<table border="0" width="852">
  <tr>
    <td width="408" height="16" bgcolor="#330099">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenário</b></font></p>
    </td>
    <td width="430" height="16" bgcolor="#330099">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#FFFFFF"><b>Relacionado com Cenários</b></font></p>
    </td>
  </tr>
<%
ANTERIOR = ""
ATUAL=""

DO UNTIL RS.EOF=TRUE

SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & UCASE(RS("CENA_CD_CENARIO_SEGUINTE")) & "'")
VALOR=TEMP("CENA_CD_CENARIO") & "-" & TEMP("CENA_TX_TITULO_CENARIO")

ATUAL=UCASE(RS("CENA_CD_CENARIO_SEGUINTE"))

IF TRIM(ATUAL)<>TRIM(ANTERIOR) THEN
	VALOR=VALOR
ELSE
	VALOR=" "
END IF

SET TEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & UCASE(RS("CENA_CD_CENARIO")) & "'")
VALOR2=TEMP2("CENA_CD_CENARIO") & "-" & UCASE(TEMP2("CENA_TX_TITULO_CENARIO"))

IF COR="WHITE" THEN
	COR="#E1E1E1"
ELSE
	COR="WHITE"
END IF

%>
  <tr>
    <td width="408" height="19"><font face="Verdana" size="2">
      <p style="margin-top: 0; margin-bottom: 0"><%=VALOR%></font></td>
    <td width="430" bgcolor="<%=COR%>" height="19"><font face="Verdana" size="2">
      <p style="margin-top: 0; margin-bottom: 0"><%=VALOR2%></font></td>
  </tr>
  <%
  TEM=1
  ANTERIOR=UCASE(RS("CENA_CD_CENARIO_SEGUINTE"))
  RS.MOVENEXT
  LOOP
  %>
</table>
<%end if%>
<%if tem=0 then%>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#800000"><b>Nenhum Registro Encontrado
para a seleção</b></font></p>
<%end if%>
</body>
</html>

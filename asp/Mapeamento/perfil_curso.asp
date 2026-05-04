<!--#include file="conecta.asp" -->
<%
set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation=3

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")
db2.CursorLocation=3

orgao = request("selOrgao")
mega = request("selMega")

if len(orgao)=2 then
	set f_orgao = db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & orgao)
	sigla_aglu = f_orgao("AGLU_SG_AGLUTINADO")
	set r_mega = db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)
	sigla_curso = r_mega("MEPR_TX_ABREVIA_CURSO")
	
	set rs = db2.execute("SELECT DISTINCT CURSO FROM [" & sigla_aglu & "] WHERE CURSO LIKE '" & sigla_curso & "%'")	

else
	set f_orgao = db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & left(orgao,2))
	sigla_aglu = f_orgao("AGLU_SG_AGLUTINADO")
	set m_orgao = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & orgao & "00000000' AND ORME_CD_STATUS='A'")
	sigla_menor = m_orgao("ORME_SG_ORG_MENOR")	
	set r_mega = db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)
	sigla_curso = r_mega("MEPR_TX_ABREVIA_CURSO")

	set rs = db2.execute("SELECT DISTINCT CURSO FROM [" & sigla_aglu & "] WHERE ORGAO='" & sigla_menor & "' AND CURSO LIKE '" & sigla_curso & "%'")
	
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Perfil Desejável para o Multiplicador</title>
</head>

<body link="#800000" vlink="#800000" alink="#800000">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="74%" id="AutoNumber2">
           <tr>
                      <td width="75%"><b><font size="2" face="Verdana" color="#000080">Perfil Desejável para Multiplicador</font></b></td>
                      <td width="25%"><p align="right"><b><font color="#800000"><a href="#" onClick="window.close()">Fechar Janela</a></font></b></td>
           </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%
do until rs.eof=true

set c = db.execute("SELECT * FROM CURSO WHERE CURS_CD_CURSO ='" & rs("curso") & "'")

set p = db2.execute("SELECT * FROM PERFIL WHERE CURSO = '" & rs("curso") & "'")

if p.eof=false then
if TRIM(p("ATUACAO"))<>"" AND TRIM(P("CONHECIMENTO"))<>"" THEN
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D7D5CC" width="69%" id="AutoNumber1" height="233">
           <tr>
                      <td width="28%" height="46" valign="top" bgcolor="#D7D5CC" bordercolor="#808080"><b><font face="Verdana" size="1">Curso</font></b></td>
                      <td width="77%" height="46" valign="top"><font size="1" face="Verdana"><B><%=C("CURS_TX_NOME_CURSO")%></B></font></td>
           </tr>
           <tr>
                      <td width="28%" height="70" valign="top" bgcolor="#D7D5CC" bordercolor="#808080"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">Área desejável </font></b></p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">de atuação </font></b></p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">do Multiplicador</font></b></td>
                      <td width="77%" height="70" valign="top"><font size="1" face="Verdana"><%=P("ATUACAO")%></font></td>
           </tr>
           <tr>
                      <td width="28%" height="115" valign="top" bgcolor="#D7D5CC" bordercolor="#808080"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">Conhecimento </font></b></p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">Desejável </font></b></p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1">do Multiplicador</font></b></td>
                      <td width="77%" height="115" valign="top"><font size="1" face="Verdana"><%=P("CONHECIMENTO")%></font></td>
           </tr>
</table>
<p>
<%
tem = tem + 1
end if
end if
rs.movenext
loop
if tem=0 then
%>
<p><b><font color="#800000">Os Cursos Selecionados não possuem Detalhamento de Perfil de Multiplicador</font></b></p>
<%
end if
%>
</body>

</html>
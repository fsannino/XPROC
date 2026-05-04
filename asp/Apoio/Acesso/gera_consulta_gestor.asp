<!--#include file="conn_consulta.asp" -->
<html>
<%
if request("excel")=1 then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

Session.LCID = 1046
chave = Ucase(request("txtchave"))

server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

set db2=server.createobject("ADODB.CONNECTION")
db2.Open "Provider=SQLOLEDB.1;server=S5200DB01\DB01;pwd=sinergiacogest;uid=usr_cogest;database=IntranetSinergia"
db2.cursorlocation=3
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Concessão de Perfil de Acesso</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<form name="frm1">
<table width="81%" height="26" border="0">
           <tr>
                      <td width="5%">
                         <div align="right">
                                   <a href="javascript:history.go(-1)"><img src="seta_esquerda_01.jpg" width="21" height="18" border="0" alt="Voltar para a Página anterior"></a></div>
                      </td>
                      <td width="5%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
                      <td width="7%">
                         <div align="right">
                                   <a href="javascript:print()"><img src="impressão.jpg" width="27" height="21" border="0" alt="Imprimir Consulta Atual">
                                   </a>
                         </div>
                      </td>
                      <td width="10%">
                      <strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong>
                      </td>
                      <td width="3%"><p align="center"></td>
                      <td width="22%">&nbsp;</td>
           </tr>
</table>
<table width="93%" border="0">
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%
on error resume next

set usuario = db.execute("SELECT USMA_TX_NOME_USUARIO AS NOME, ORME_CD_ORG_MENOR FROM USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & CHAVE & "'")
set orgao = db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & usuario("ORME_CD_ORG_MENOR") & "'")

if err.number=0 then
err.clear
%>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#000080">Chave Consultada : <b><%=chave%> - <%=usuario("nome")%> - <%=orgao("lotacao")%></b></font></p>
<hr>
<%
ssql=""
ssql="SELECT * FROM dbo.USMA_MICRO_R3_VISAO_R3 "
ssql=ssql + " WHERE (MCPR_NR_SEQ_MACRO_PERFIL=98) AND (MIPE_NR_SEQ_MICRO_PERFIL=1) AND(USMA_CD_USUARIO = '" & chave & "')"

set perfilr3 = db.execute(ssql)

if perfilr3.eof=false then
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="76%" id="AutoNumber2" height="19">
           <tr>
                      <td width="1%" height="19">&nbsp;</td>
                      <td width="23%" bgcolor="#D7D5CC" height="19"><font face="Verdana" size="2" color="#000080">Função de Negócio :</font></td>
                      <td width="121%" bgcolor="#D7D5CC" height="19"><font face="Verdana" size="2" color="#000080">&nbsp;<b>HR.03 - GESTOR DE PESSOAS</b></font></td>
           </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="61%" id="AutoNumber1">
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="73%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Perfil no R/3</font></b></td>
                      <td width="18%" bgcolor="#D7D5CC">&nbsp;</td>
           </tr>
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="91%" colspan="2"><font face="Verdana" size="2" color="#000080">Z:HR_PB001_GESTOR_PESSOAS</font></td>
           </tr>
<%
tem = tem + 1
end if
%>          
<b>
<%
if tem=0 then
%>
<font color="#800000"><b>Nenhum Registro Encontrado para a Seleção</b></font>
<%
end if
%>
</font></b>
<%else%>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#000080">Usuário não encontrado</b></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<hr>
<%end if%>
</form>
</body>

</html>
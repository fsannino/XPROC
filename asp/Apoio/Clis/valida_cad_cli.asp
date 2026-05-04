<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conn_consulta.asp" -->
<%
server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

chave=request("txtchave")
edita=request("txtedita")

if edita=0 then
	ssql=""
	ssql="INSERT INTO " & Session("Perfixo") & "CLI "
	ssql=ssql+"(USMA_CD_USUARIO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
	ssql=ssql+"VALUES('" & chave & "', "
	ssql=ssql+"'I', "
	ssql=ssql+"'" & Session("cdUsuario") & "', "
	ssql=ssql+"GETDATE())"
	
	oper="I"
else
	ssql=""
	ssql="UPDATE " & Session("Perfixo") & "CLI "
	ssql=ssql+"SET ATUA_TX_OPERACAO='A', "
	ssql=ssql+"ATUA_CD_NR_USUARIO='" & Session("cdUsuario") & "', "
	ssql=ssql+"ATUA_DT_ATUALIZACAO=GETDATE()"
	ssql=ssql+" WHERE USMA_CD_USUARIO='" & chave & "'"

	oper="A"
end if

db.execute(ssql)

ssql=""
ssql="INSERT INTO LOG_APOIO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
ssql=ssql+" VALUES('" & chave & "', "
ssql=ssql+"3, "
ssql=ssql+"'" & oper & "','" & Session("CdUsuario") & "', GETDATE()) "

db.execute(ssql)
%>

<html>
<head>

<title>Base de Dados de Coordenadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function envia()
{
window.location = "cad_orgao_cli.asp?chave=" + this.chave.value
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onLoad="envia()">
<p>
<input type="hidden" name="edita" size="11" value="<%=edita%>">
<input type="hidden" name="chave" size="11" value="<%=chave%>">
</p>
<p>&nbsp;&nbsp;&nbsp; <font color="#000080">Carregando Orgãos...por favor,
aguarde...</font></p>
</body>
</html>

<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->
<%
server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

chave=request("chave")
tipo=request("tipo")

SSQL1="DELETE FROM APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & tipo
SSQL2="DELETE FROM APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & tipo
SSQL3="DELETE FROM APOIO_LOCAL_MODULO WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & tipo
SSQL4="DELETE FROM APOIO_LOCAL_ONDA WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & tipo
SSQL5="DELETE FROM APOIO_LOCAL_MULT WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & tipo

db.execute(ssql1)
db.execute(ssql2)
db.execute(ssql3)
db.execute(ssql4)
db.execute(ssql5)

ssql=""
ssql="INSERT INTO LOG_APOIO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
ssql=ssql+" VALUES('" & chave & "', "
ssql=ssql+"" & tipo & ", "
ssql=ssql+"'E','" & Session("CdUsuario") & "', GETDATE()) "

db.execute(ssql)
%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio...Redirecionando...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
<p>
<input type="hidden" name="edita" size="11" value="<%=edita%>">
<input type="hidden" name="chave" size="11" value="<%=chave%>">
<input type="hidden" name="atrib" size="11" value="<%=atrib%>">
</p>
</body>
</html>

<script>
alert('Associação do Registro Excluída com Sucesso!')
window.location="menu.asp"
</script>

<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->
<%
server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

chave=request("txtchave")
atrib=request("sel_Apoio")
orgao=request("txtorgao")
momento=request("txtmomento")
obs=request("str_obs")
situacao=request("str_Ativo")
relacao=request("txtvinculo")

modulo=request("txtmodulo")
onda=request("txtonda")

edita=request("txtedita")

if edita=0 then
	ssql=""
	ssql="INSERT INTO " & Session("Perfixo") & "APOIO_LOCAL_MULT "
	ssql=ssql+"(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APLO_NR_MOMENTO,APLO_TX_OBS,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,APLO_NR_SITUACAO,APLO_NR_RELACAO_EMPREGO) "
	ssql=ssql+"VALUES('" & chave & "', "
	ssql=ssql+"" & atrib & ", "
	ssql=ssql+"'" & orgao & "', "
	if momento <> "" then
       ssql=ssql+"" & momento & ", "
	else
       ssql=ssql+" null , "	
	end if   
	ssql=ssql+"'" & obs & "', "
	ssql=ssql+"'I', "
	ssql=ssql+"'" & Session("cdUsuario") & "', "
	ssql=ssql+"GETDATE(), "
	ssql=ssql+"" & situacao & ", "
	ssql=ssql+"'" & relacao & "') "
	
	oper="I"
else
	ssql=""
	ssql="UPDATE " & Session("Perfixo") & "APOIO_LOCAL_MULT "
	ssql=ssql+"SET ORME_CD_ORG_MENOR='" & orgao & "', "
	if momento <> "" then
	   ssql=ssql+"APLO_NR_MOMENTO=" & momento & ", "
	end if   
	ssql=ssql+"APLO_TX_OBS='" & obs & "', "
	ssql=ssql+"ATUA_TX_OPERACAO='A', "
	ssql=ssql+"ATUA_CD_NR_USUARIO='" & Session("cdUsuario") & "', "
	ssql=ssql+"ATUA_DT_ATUALIZACAO=GETDATE(), "
	ssql=ssql+"APLO_NR_SITUACAO=" & situacao & ", "
	ssql=ssql+"APLO_NR_RELACAO_EMPREGO='" & relacao & "' "
	ssql=ssql+" WHERE USMA_CD_USUARIO='" & chave & "' AND APLO_NR_ATRIBUICAO=" & atrib
	
	oper="A"
end if

db.execute(ssql)

ssql=""
ssql="INSERT INTO LOG_APOIO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
ssql=ssql+" VALUES('" & chave & "', "
ssql=ssql+"" & atrib & ", "
ssql=ssql+"'" & oper & "','" & Session("CdUsuario") & "', GETDATE()) "

db.execute(ssql)


Sub Grava_modulo(SChave, SAtribu, SModulo)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_MODULO "
	ssql=ssql+"VALUES('" & SChave & "',"
	ssql=ssql+"" & SAtribu & ","
	ssql=ssql+"" & SModulo & ","
	ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

	db.execute(ssql)
	
end sub

Sub Grava_onda(SChave, SAtribu, SOnda)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ONDA "
	ssql=ssql+"VALUES('" & SChave & "',"
	ssql=ssql+"" & SAtribu & ","
	ssql=ssql+"" & SOnda & ","
	ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

	db.execute(ssql)
	
end sub

a = "DELETE FROM " & Session("Prefixo") & "APOIO_LOCAL_MODULO WHERE APLO_NR_ATRIBUICAO=" & ATRIB & " AND USMA_CD_USUARIO='" & CHAVE & "'"
b = "DELETE FROM " & Session("Prefixo") & "APOIO_LOCAL_ONDA WHERE APLO_NR_ATRIBUICAO=" & ATRIB & " AND USMA_CD_USUARIO='" & CHAVE & "'"

db.execute(a)
db.execute(b)

str_valor = modulo

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if
tamanho = Len(str_valor)
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If
tamanho = Len(str_valor)
contador = 1
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
			call Grava_modulo(chave,atrib,str_atual)
	   	
			valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

str_valor = onda

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if
tamanho = Len(str_valor)
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If
tamanho = Len(str_valor)
contador = 1
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
			call Grava_onda(chave,atrib,str_atual)
	   	
			valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop
%>

<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio...Redirecionando...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function envia()
{
window.location = "cad_orgao.asp?chave=" + this.chave.value + "&atribb=" + this.atrib.value
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onLoad="envia()">
<p>
<input type="hidden" name="edita" size="11" value="<%=edita%>">
<input type="hidden" name="chave" size="11" value="<%=chave%>">
<input type="hidden" name="atrib" size="11" value="<%=atrib%>">
</p>
<p>&nbsp;&nbsp;&nbsp; <font color="#000080">Carregando Orgãos...por favor,
aguarde...</font></p>
</body>
</html>
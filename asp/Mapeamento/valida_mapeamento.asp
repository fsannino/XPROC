<!--#include file="conecta.asp" -->
<%
chave = request("selFunc")
mega = request("selMega")

tcurso=0

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

b = "SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega
set temp = db.execute(b)

a = "DELETE FROM " & Session("Prefixo") & "APOIO_LOCAL_CURSO WHERE CURS_CD_CURSO LIKE '" & temp("MEPR_TX_ABREVIA_CURSO") & "%' AND USMA_CD_USUARIO='" & CHAVE & "'"
db.execute(a)

set objUSR = server.createobject("Seseg.Usuario")

if objUSR.GetUsuario then
	resp=objUSR.sei_chave
	nome=objUSR.sei_nome
else
	response.redirect "erro.asp?op=3"
end if

set tem_mult = db.execute("SELECT * FROM APOIO_LOCAL_MULT WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")

if tem_mult.eof=true and len(request("txtcurso"))>2 then

	set lot = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR AS LOTACAO, USMA_CD_VINCULO AS VINCULO FROM USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & chave & "'")
	
	if lot("vinculo")="F" THEN
		vinculo="E"
	else
		vinculo="C"
	end if
	
	ssql=""
	ssql="INSERT INTO APOIO_LOCAL_MULT(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ORME_CD_ORG_MENOR, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, APLO_NR_SITUACAO, APLO_NR_RELACAO_EMPREGO)"
	ssql=ssql+" VALUES('" & chave & "',2 , '" & lot("lotacao") & "', "
	ssql=ssql+"'A','" & resp & "', GETDATE(), 1, '" & vinculo & "')"
	
	db.execute(ssql)

	ssql=""
	ssql="INSERT INTO APOIO_LOCAL_ONDA(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ONDA_CD_ONDA, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
	ssql=ssql+" VALUES('" & chave & "',2 , 6, "
	ssql=ssql+"'A','" & resp & "', GETDATE())"
	
	db.execute(ssql)
	
	org_mult = left(lot("lotacao"),7)
	
	set org_m = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR LIKE '" & org_mult & "%' AND ORME_CD_STATUS='A'")
	
	do until org_m.eof=true
		ssql=""
		ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
		ssql=ssql+"VALUES('" & Chave & "',"
		ssql=ssql+"2 ,"
		ssql=ssql+"'" & org_m("ORME_CD_ORG_MENOR") & "',"
		ssql=ssql+"'',"
		ssql=ssql+"'I','" & resp & "',GETDATE())"
		
		on error resume next
		db.execute(ssql)
		err.clear
	
		org_m.movenext
	loop

end if

ssql=""
ssql="INSERT INTO LOG_APOIO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
ssql=ssql+" VALUES('" & chave & "',4 , "
ssql=ssql+"'A','" & resp & "', GETDATE()) "

db.execute(ssql)

Sub Grava_Curso(SChave, SMega, sCurso)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_CURSO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, CURS_CD_CURSO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
	ssql=ssql+"VALUES('" & SChave & "', 2, "
	ssql=ssql+"'" & SCurso & "',"	
	ssql=ssql+"'I','" & resp & "',GETDATE())"

	db.execute(ssql)
	
	tcurso = tcurso + 1
	
end sub

str_valor = request("txtcurso")

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
        
			call Grava_Curso(chave,mega,str_atual)
	   	
			valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

mensagem="O Mapeamento foi efetuado com Sucesso!"

if tcurso=0 then

	set tem_curso = db.execute("SELECT * FROM APOIO_LOCAL_CURSO WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")
	
	if tem_curso.eof=true then
		db.execute("DELETE FROM APOIO_LOCAL_ONDA WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")
		db.execute("DELETE FROM APOIO_LOCAL_ORGAO WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")
		db.execute("DELETE FROM APOIO_LOCAL_MULT WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")				
	end if

end if

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form name="frm1">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="487">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="418" valign="top"><img border="0" src="lado.jpg" width="83" height="417"></td>
                      <td width="87%" height="418" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3" height="199">
                         <tr>
                                    <td width="100%" height="106" colspan="4" align="center"><b><font face="Verdana" color="#800000"><%=mensagem%></font></b></td>
                         </tr>
                         <tr>
                                    <td width="25%" height="19" align="center">&nbsp;</td>
                                    <td width="9%" height="19" align="center">&nbsp;</td>
                                    <td width="41%" height="19" align="center">&nbsp;</td>
                                    <td width="25%" height="19" align="center">&nbsp;</td>
                         </tr>
                         <tr>
                                    <td width="25%" height="36" align="center">&nbsp;</td>
                                    <td width="9%" height="36" align="center"><b><font face="Verdana"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="41%" height="36" align="left"><b><font face="Verdana" size="2"><a href="menu.asp">Retornar ao Menu Principal</a></font></b></td>
                                    <td width="25%" height="36" align="center">&nbsp;</td>
                         </tr>
                         <tr>
                                    <td width="25%" height="35" align="center">&nbsp;</td>
                                    <td width="9%" height="35" align="center"><b><font face="Verdana"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="41%" height="35" align="left"><b><font face="Verdana" size="2"><a href="mapeamento.asp">Mapear outro Multiplicador</a></font></b></td>
                                    <td width="25%" height="35" align="center">&nbsp;</td>
                         </tr>
                         <tr>
                                    <td width="25%" height="19">&nbsp;</td>
                                    <td width="9%" height="19">&nbsp;</td>
                                    <td width="41%" height="19">&nbsp;</td>
                                    <td width="25%" height="19">&nbsp;</td>
                         </tr>
                         </table>
                      </td>
           </tr>
</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>
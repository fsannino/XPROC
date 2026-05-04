<%
funcao=request("funcao")
transacao=request("transac")
opt=request("opt")
user=request("user")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'Altera a situação da Transação na tabela MACRO_PERFIL_TRANSACAO para '2'
conn_db.execute("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO SET MCPT_NR_SITUACAO_ALTERACAO_FUNC=2 WHERE TRAN_CD_TRANSACAO='" & Transacao & "'")

'Recupera todos os MACRO-PERFIL associados àquela transacao
set verifica=conn_db.execute("SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL FROM " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO WHERE TRAN_CD_TRANSACAO='" & request("Transac") & "'")

DO UNTIL VERIFICA.EOF=TRUE

	set correio=server.createobject("Persits.MailSender")
	correio.host = "harpia.petrobras.com.br"

	set temp=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & verifica("MCPR_NR_SEQ_MACRO_PERFIL"))
	
	macro=temp("MCPE_TX_NOME_TECNICO")

	resp=temp("ATUA_CD_NR_USUARIO")
	
	origem = "x-proc@admin.com"
	
	set quem=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE ATUA_CD_NR_USUARIO='" & temp("ATUA_CD_NR_USUARIO") & "'")

	destino=""
	
	if quem.eof=false then
		destino = quem("USUA_TX_EMAIL_EXTERNO")
	end if
	
	if isnull(destino) then
		destino=""
	end if
	
	assunto = "ALTERAÇÃO DE MACRO PERFIL / FUNÇAO DE NEGÓCIO / ALTERADA EM : " & now & " / ALTERADA POR :" & request("chave")
	mensagem = "A função " & funcao & " utilizada para contrução do macro perfil " & macro & " foi alterada com a " & opt & " da transação " & transacao & " ."
	
	correio.from = origem
	correio.AddAddress destino
	correio.Subject = assunto
	correio.body = mensagem
		
	'Envia correio para o responsável do MACRO-PERFIL corrente
	if destino<>"" then
		correio.send
	end if
	
	set correio=nothing
	
	'Atualiza o Status do MACRO-PERFIL para 'AT'
	conn_db.execute("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='AT' WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & verifica("MCPR_NR_SEQ_MACRO_PERFIL"))
	
	SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & verifica("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
	ATUAL=HIST("CODIGO")
   	ATUAL = ATUAL + 1
	if atual > 1 then
		atual = atual
   	else
   		atual=1
	end if

	SSQL=""
	SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	SSQL=SSQL+"VALUES(" & request("MEGA") & ",'EXCLUSÃO DE TRANSAÇÃO', " & ATUAL &", " & verifica("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'AT', 'I', '" & request("User") & "', GETDATE())"
        		
	conn_db.execute(ssql)
	
	VERIFICA.MOVENEXT
LOOP
%>

<html>
<head>
<title>Nova pagina 1</title>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" onload="javascript:window.close()">
  <p align="center">&nbsp;</p>
  <p align="center"><b><font size="3" face="Verdana" color="#330099"></font></b></p>
</body>
</html>
<%
funcao=request("funcao")
transacao=request("transac")
opt=request("opt")
user=request("user")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'Verifica todos os MACRO-PERFIL relacionados àquela função
str_SQL = ""
str_SQL = str_SQL & " SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL, MCPE_TX_SITUACAO FROM " & Session("PREFIXO") & "MACRO_PERFIL "
str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO='" & request("funcao") & "'"

set ver_macros = conn_db.execute(str_SQL)

do until ver_macros.eof=true

   str_Texto_Historico1 = ""
   str_Texto_Historico2 = ""
   str_Proximo_Status = ver_macros("MCPE_TX_SITUACAO")

   'Vê se a sequência MACRO-PERFIL / TRANSAÇÃO existe
   
   str_SQL = ""
   str_SQL = str_SQL & " SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL FROM " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") 
   str_SQL = str_SQL & " AND TRAN_CD_TRANSACAO='" & request("Transac") & "'"
   
   set verifica=conn_db.execute(str_SQL)
   
   if verifica.eof=true then	

	  'Se o registro não existir, inseri-lo	
	  str_SQL = ""
	  str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO( "
	  str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO "
	  str_SQL = str_SQL & " , MCPT_NR_SITUACAO_ALTERACAO, MCPT_NR_SITUACAO_ALTERACAO1, MCPT_NR_SITUACAO_ALTERACAO_FUNC, MCPT_NR_SITUACAO_PROCESSAMENTO "
	  str_SQL = str_SQL & " , ATUA_TX_OPERACAO "
	  str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	  str_SQL = str_SQL & " VALUES (" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & TRANSACAO & "', " & request("mega") & ", "
	  str_SQL = str_SQL & " 0, 0, 1, 0, "
	  str_SQL = str_SQL & " 'I', '" & user &  "', GETDATE())"	

	  'response.write str_sql	
	  conn_db.execute(str_SQL)
	  
	  str_SQL = ""
	  str_SQL = str_SQL & " SELECT  MEPR_TX_DESC_MEGA_PROCESSO "
      str_SQL = str_SQL & " FROM    dbo.MEGA_PROCESSO "
      str_SQL = str_SQL & " WHERE   MEPR_CD_MEGA_PROCESSO = " & request("mega")
	  'response.write str_sql	
	  set rds_Ds_Mega = conn_db.execute(str_SQL)
	  if not rds_Ds_Mega.EOF then
	     str_Ds_Mega = rds_Ds_Mega("MEPR_TX_DESC_MEGA_PROCESSO")
	  else
	     str_Ds_Mega = ""
	  end if
	  rds_Ds_Mega.close
	  set rds_Ds_Mega = Nothing
      str_Texto_Historico1 = "INCLUIDO A TRANSAÇÃO " & TRANSACAO & " PELO MEGA PROCESSO " & str_Ds_Mega 
      
	  'Gravar objetos referentes à transacao
	  ' rotina retirada da LOGICA - JOAO LUIZ
	  'str_SQL = ""
	  'str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO( "
	  'str_SQL = str_SQL & " MPTO_TX_SIT_ALTERACAO_VALOR, MPTO_TX_SIT_ALTERACAO_VALOR1, MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, "
	  'str_SQL = str_SQL & " TROB_TX_CAMPO, TROB_TX_OBJETO, MPTO_TX_VALORES, TROB_TX_CRITICO , ATUA_TX_OPERACAO, "
	  'str_SQL = str_SQL & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
	  'str_SQL = str_SQL & " (SELECT '0','0', " & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") & ", TRAN_CD_TRANSACAO, TROB_TX_CAMPO, "
	  'str_SQL = str_SQL & " TROB_TX_OBJETO, TRON_TX_VALORES, TROB_TX_CRITICO, "
	  'str_SQL = str_SQL & " 'I', '" & user &  "', GETDATE() "
	  'str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_OBJETO "
	  'str_SQL = str_SQL & " WHERE TRAN_CD_TRANSACAO = '" & Transacao & "')"	
	  'conn_db.execute(str_SQL)

	  str_Necessita_Aprovacao = 0	
      '===========================================================================
      ' CRIAR A ROTINA DE VERIFICAÇÃO SE ESTA TRANSAÇÃO NECESSITA DE VALIDAÇÃO
	  ' NÃ DESENVOLVIDA COM ORIENTAÇÃO DA KATIA E MARCELO
	  '===========================================================================
	
	  set correio=server.createobject("Persits.MailSender")
	  correio.host = "harpia.petrobras.com.br"
	  set temp=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))	
	  macro=temp("MCPE_TX_NOME_TECNICO")	
	  set quem=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE ATUA_CD_NR_USUARIO='" & temp("ATUA_CD_NR_USUARIO") & "'")	
	  destino=""	
	  if quem.eof=false then
		 destino = quem("USUA_TX_EMAIL_EXTERNO")
	  end if	
	  mensagem = "A função " & funcao & " utilizada para contrução do macro perfil " & macro & " foi alterada com a " & opt & " da transação " & transacao & " ."
	  de = "x-proc@admin.com"
	  para = destino
	  assunto = "ALTERAÇÃO DE MACRO PERFIL / FUNÇAO DE NEGÓCIO"
	  if destino<>"" then
		 correio.From = de
		 correio.AddAddress para
		 correio.subject = assunto
		 correio.body = mensagem
		 correio.send
	  end if
	  
      if str_Necessita_Aprovacao = 0 then
	     if ver_macros("MCPE_TX_SITUACAO") = "EE" OR ver_macros("MCPE_TX_SITUACAO") = "EC" OR ver_macros("MCPE_TX_SITUACAO") = "AR" then
	        ' EM ELABORAÇÃO  - EM CRIAÇÃO - EM ALTERAÇÃO
			' NÃO FAZ NADA
	     elseif ver_macros("MCPE_TX_SITUACAO") = "CR" OR ver_macros("MCPE_TX_SITUACAO") = "AP" then
	        ' CRIADO - ALTERADO NO R3
 	        str_Proximo_Status = "AR"
		    str_Texto_Historico2 = "MUDADO O STATUS PARA EM ALTERAÇÃO NO R3"
	     elseif ver_macros("MCPE_TX_SITUACAO") = "EA" OR ver_macros("MCPE_TX_SITUACAO") = "NA" OR ver_macros("MCPE_TX_SITUACAO") = "RE" then
	        ' EM APROVAÇÃO - NÃO APROVADO - RECUSADO NO R3
			' NÃO FAZ NADA
	     elseif ver_macros("MCPE_TX_SITUACAO") = "EX" OR ver_macros("MCPE_TX_SITUACAO") = "MR" OR ver_macros("MCPE_TX_SITUACAO") = "EL" then
	        ' EXCLUIDO A FUNÇÃO - MUDADO PARA REFERENCIA - EXCLUÍDO NO X PROC - EXCLUÍDO NO R3
            ' NÃO FAZ NADA
		 elseif ver_macros("MCPE_TX_SITUACAO") = "ER" OR ver_macros("MCPE_TX_SITUACAO") = "EP" then
	        ' EM EXCLUSÃO NO R3 - EXCLUÍDO NO R3
			' NÃO FAZ NADA
	     end if
	  else
	     if ver_macros("MCPE_TX_SITUACAO") <> "EA" then
	 	    str_Proximo_Status = "EA"
	        str_Texto_Historico2 = "MUDADO O STATUS PARA EM VALIDAÇÃO"		 
		 end if	
	  end if
	  if str_Proximo_Status <> "" then
	  	 str_SQL = ""
	     str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='" & str_Proximo_Status & "'"
		 str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL")
	     conn_db.execute(str_SQL)
      end if	  
	  	  
	  if str_Texto_Historico1 <> "" then
	     SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))        		
	     ATUAL=HIST("CODIGO")
   	     ATUAL = ATUAL + 1
	     if atual > 1 then
		    atual = atual
   	     else
   		    atual=1
	     end if
	     SSQL=""
	     SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	     SSQL=SSQL+"VALUES(" & request("MEGA") & ",'" &  str_Texto_Historico1 & "', " & ATUAL &", " & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & str_Proximo_Status & "', 'I', '" & request("User") & "', GETDATE())"        		
	     conn_db.execute(ssql)		
      end if	

	  if str_Texto_Historico2 <> "" then
	     SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))        		
	     ATUAL=HIST("CODIGO")
   	     ATUAL = ATUAL + 1
	     if atual > 1 then
		    atual = atual
   	     else
   		    atual=1
	     end if
	     SSQL=""
	     SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	     SSQL=SSQL+"VALUES(" & request("MEGA") & ",'" &  str_Texto_Historico2 & "', " & ATUAL &", " & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & str_Proximo_Status & "', 'I', '" & request("User") & "', GETDATE())"        		
	     conn_db.execute(ssql)		
      end if	
	  	  	 
   else

      str_SQL = ""
      str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
	  str_SQL = str_SQL & " SET MCPT_NR_SITUACAO_ALTERACAO_FUNC = 1 ,"
	  str_SQL = str_SQL & " MCPT_NR_SITUACAO_PROCESSAMENTO = 0 ,"	   
      str_SQL = str_SQL & " ATUA_CD_NR_USUARIO =  '" & Session("CdUsuario") & "',"
      str_SQL = str_SQL & " ATUA_TX_OPERACAO = 'A', "
      str_SQL = str_SQL & " ATUA_DT_ATUALIZACAO = GETDATE() "	   	   	   
	  str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") 
	  str_SQL = str_SQL & " AND TRAN_CD_TRANSACAO='" & Transacao & "'"
	  
	  conn_db.execute(str_SQL)
	  
	  set correio=server.createobject("Persits.MailSender")
	  correio.host = "harpia.petrobras.com.br"
	  set temp=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))	
	  macro=temp("MCPE_TX_NOME_TECNICO")	
	  set quem=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE ATUA_CD_NR_USUARIO='" & temp("ATUA_CD_NR_USUARIO") & "'")
	  destino=""	
	  if quem.eof=false then
		 destino = quem("USUA_TX_EMAIL_EXTERNO")
	  end if	
	  mensagem = "A função " & funcao & " utilizada para contrução do macro perfil " & macro & " foi alterada com a " & opt & " da transação " & transacao & " ."
	  de = "x-proc@admin.com"
	  para = destino
	  assunto = "ALTERAÇÃO DE MACRO PERFIL / FUNÇAO DE NEGÓCIO"
	  if destino<>"" then
		 correio.From = de
		 correio.AddAddress para
		 correio.subject = assunto
		 correio.body = mensagem
		 correio.send
	  end if	
	  conn_db.execute("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='" & str_Proximo_Status & "' WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))
	  SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL"))        		
	  ATUAL=HIST("CODIGO")
   	  ATUAL = ATUAL + 1
	  if atual > 1 then
		 atual = atual
   	  else
   		 atual=1
	  end if
	  SSQL=""
	  SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	  SSQL=SSQL+"VALUES(" & request("MEGA") & ",'INCLUSÃO DE TRANSAÇÃO', " & ATUAL &", " & ver_macros("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & str_Proximo_Status & "', 'I', '" & request("User") & "', GETDATE())"        		
	  conn_db.execute(ssql)

   end if

   ver_macros.movenext

loop

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
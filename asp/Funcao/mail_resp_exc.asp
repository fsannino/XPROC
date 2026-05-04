<%
funcao=request("funcao")
transacao=request("transac")
opt=request("opt")
user=request("user")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'==================================================================================================
' começa aqui
'==================================================================================================
'Recupera todos os MACRO-PERFIL associados àquela FUNCAO

str_SQL = ""
str_SQL = str_SQL & " SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL, MCPE_TX_NOME_TECNICO, MCPE_TX_SITUACAO, ATUA_CD_NR_USUARIO "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL "
str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'"

set rsdMacroPerfil = conn_db.execute(str_SQL)

Do While not rsdMacroPerfil.EOF
   str_Proximo_Status = rsdMacroPerfil("MCPE_TX_SITUACAO")
'=====================================================================================================
' VERIFICA SE AQUELE MACRO JÁ FOI CRIADO NO R3
    str_SQL = ""
	str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL "
	str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO "
	str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
	str_SQL = str_SQL & " and MHVA_TX_SITUACAO_MACRO = 'CR' " 
	set rds_Temp = conn_db.execute(str_SQL)

	if rds_Temp.EOF then
	   str_Ja_Criado = 0
	else
	   str_Ja_Criado = 1
	end if
	
	rds_Temp.close
	set rds_Temp = Nothing
'======================================================================================================	
	str_Deleta_Tudo = 0
    str_Texto_Historico2 = ""
    str_Texto_Historico1 = " EXCLUSÃO DE TRANSAÇÃO - " & transacao
	str_Deleta_Status_Transacao = 0
'======================================================================================================				   
    if rsdMacroPerfil("MCPE_TX_SITUACAO") = "EE" or rsdMacroPerfil("MCPE_TX_SITUACAO") = "EC" or rsdMacroPerfil("MCPE_TX_SITUACAO") = "ER" then
	   ' SE EM ELABORAÇÃO - EM CRIAÇÃO - EM EXCLUSÃO - RECUSADO
	   str_Deleta_Tudo = 1
	elseif rsdMacroPerfil("MCPE_TX_SITUACAO") = "EX" OR rsdMacroPerfil("MCPE_TX_SITUACAO") = "MR" OR rsdMacroPerfil("MCPE_TX_SITUACAO") = "EL" OR rsdMacroPerfil("MCPE_TX_SITUACAO") = "EP" then   
	   ' EXCLUIDO A FUNÇÃO - MUDADO A FUNÃO PARA REFERENCIA - EXCLUIDO NO XPROC - EXCLUIDO NO R3
	   ' NÃO FAZ NADA
	   str_Texto_Historico1 = ""
    elseif rsdMacroPerfil("MCPE_TX_SITUACAO") = "EA" OR rsdMacroPerfil("MCPE_TX_SITUACAO") = "NA" then
       ' EM APROVAÇÃO - NÃO APROVADO
	   str_SQL = ""
	   str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL "
	   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA "
	   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
	   str_SQL = str_SQL & " and TRAN_CD_TRANSACAO <> '" & transacao & "'"
	   str_SQL = str_SQL & " and MAOA_TX_AUTORIZADO <> '1' AND MAOA_TX_AUTORIZADO <> '3'"
	   set rds_Temp = conn_db.execute(str_SQL)
	   if rds_Temp.EOF then
          str_Altera_Status_Macro = 1
		  str_Texto_Historico2 = "ENVIADO PARA CRIAÇÃO NO R3 EM FUNÇÃO DE EXCLUSÃO DE TRANSAÇÃO"
		  str_Proximo_Status = "EC"		     
       end iF
	   rds_Temp.close
	   set rds_Temp = nothing
       str_Deleta_Tudo = 1
    elseif rsdMacroPerfil("MCPE_TX_SITUACAO") = "CR" OR rsdMacroPerfil("MCPE_TX_SITUACAO") = "AP" then
       ' SE CRIADO NO R3  - ALTERADO NO R3
	   str_Altera_Status_Transacao = 1  
       str_Altera_Status_Macro = 1
	   str_Proximo_Status = "AR"
    elseif rsdMacroPerfil("MCPE_TX_SITUACAO") = "AR"  then
       ' SE EM ALTERAÇÃO NO R3
	   str_Altera_Status_Transacao = 1  
	elseif rsdMacroPerfil("MCPE_TX_SITUACAO") = "RE" then
       if str_Ja_Criado = 1 then
          str_Altera_Status_Transacao = 1
       else
          str_Deleta_Tudo = 1      
	   end if		  		 	   
	end if
    if str_Proximo_Status = "AR" then
	   str_Ds_Proximo_Status = "EM ALTERAÇÃO "
	elseif str_Proximo_Status = "EC" then
	   str_Ds_Proximo_Status = "EM CRIAÇÃO "	
    else
	   str_Ds_Proximo_Status = rsdMacroPerfil("MCPE_TX_SITUACAO")
	end if
	if str_Altera_Status_Macro = 1 then
       str_SQL = ""
	   str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL "
	   str_SQL = str_SQL & " SET MCPE_TX_SITUACAO='" & str_Proximo_Status & "'"
	   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
	   conn_db.execute(str_SQL)
	   str_Texto_Historico2 = "ENVIADO PARA " & str_Ds_Proximo_Status & "NO R3 EM FUNÇÃO DE EXCLUSÃO DE TRANSAÇÃO"		     
    end iF
    
	if str_Altera_Status_Transacao = 1 then
       'Altera a SITUACAO da Transação na tabela MACRO_PERFIL_TRANSACAO para '2' INDICANDO QUE AQUELA TRANSAÇÃO FOI EXCLUÍDA
       str_SQL = "" 
       str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO SET "
       str_SQL = str_SQL & " MCPT_NR_SITUACAO_ALTERACAO_FUNC = 2 ,"
       str_SQL = str_SQL & " MCPT_NR_SITUACAO_PROCESSAMENTO = 0 ,"	   
       str_SQL = str_SQL & " ATUA_CD_NR_USUARIO =  '" & Session("CdUsuario") & "',"
       str_SQL = str_SQL & " ATUA_TX_OPERACAO = 'A', "
       str_SQL = str_SQL & " ATUA_DT_ATUALIZACAO = GETDATE() "	   	   	   
       str_SQL = str_SQL & " WHERE TRAN_CD_TRANSACAO='" & transacao & "'"
       str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL = " & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
       conn_db.execute(str_SQL)	 	
	end if
	if str_Deleta_Tudo = 1 then
       str_SQL = ""
       str_SQL = str_SQL & " SELECT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
	   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & funcao & "'"
       str_SQL = str_SQL & " and  TRAN_CD_TRANSACAO = '" & transacao & "'"
	   'RESPONSE.Write(str_SQL)
	   set rds_Repete_Transacao = conn_db.execute(str_SQL)
	   'RESP = 0
	   if rds_Repete_Transacao.EOF then	   	      	   	   
	      'RESP = 1
          str_SQL = ""
          str_SQL = str_SQL & " DELETE FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA "
          str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
          str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & transacao & "'"
          conn_db.execute(str_SQL)

          str_SQl = ""
          str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO " 
          str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
          str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & transacao & "'"	   
          conn_db.execute(str_SQl)    
	   		
          str_SQL = ""
          str_SQL = str_SQL & " DELETE FROM " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
          str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
          str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & transacao & "'"
          conn_db.execute(str_SQL)
	   end if
	   'RESPONSE.Write(resp)
	   'RESPONSE.Write(resp)
	   rds_Repete_Transacao.close
	   set rds_Repete_Transacao = Nothing	  
    end if

    if str_Texto_Historico1 <> "" then
	   SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL"))        		
	   ATUAL=HIST("CODIGO")
   	   ATUAL = ATUAL + 1
	   if atual > 1 then
		  atual = atual
   	   else
   		  atual=1
	   end if
	   SSQL=""
	   SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	   SSQL=SSQL+"VALUES(" & request("MEGA") & ",'"  & str_Texto_Historico1 & "', " & ATUAL & ", " & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & str_Proximo_Status & "', 'I', '" & request("User") & "', GETDATE())"        		
	   conn_db.execute(ssql)
	   HIST.close
	   set HIST = Nothing	  
       Call Passa_Email(rsdMacroPerfil("MCPE_TX_NOME_TECNICO"),str_Proximo_Status,rsdMacroPerfil("ATUA_CD_NR_USUARIO")) 
	end if   
	
    if str_Texto_Historico2 <> "" then
	   SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL"))        		
	   ATUAL=HIST("CODIGO")
   	   ATUAL = ATUAL + 1
	   if atual > 1 then
		  atual = atual
   	   else
   		  atual=1
	   end if
	   SSQL=""
	   SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	   SSQL=SSQL+"VALUES(" & request("MEGA") & ",'"  & str_Texto_Historico2 & "', " & ATUAL & ", " & rsdMacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & str_Proximo_Status & "', 'I', '" & request("User") & "', GETDATE())"        		
	   conn_db.execute(ssql)
	   HIST.close
	   set HIST = Nothing
       Call Passa_Email(rsdMacroPerfil("MCPE_TX_NOME_TECNICO"),str_Proximo_Status,rsdMacroPerfil("ATUA_CD_NR_USUARIO")) 	   
	end if   
    
	rsdMacroPerfil.MOVENEXT
	
LOOP

Sub Passa_Email(macro,proximo_status,responsavel)
	set correio=server.createobject("Persits.MailSender")
	correio.host = "harpia.petrobras.com.br"
	origem = "x-proc@admin.com"		
    'VERIFICA QUEM CRIOU O MACRO PERFIL PARA SER ENCAMINHADO UM EMAIL		
	set quem=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE ATUA_CD_NR_USUARIO='" & responsavel & "'")	
	if quem.eof=false then
	   if not IsNull(quem("USUA_TX_EMAIL_EXTERNO")) then
	      destino = quem("USUA_TX_EMAIL_EXTERNO")
	   else
	      destino=""
	   end if 	   
	else
	   destino=""		
	end if		
	assunto = "ALTERAÇÃO DE MACRO PERFIL/FUNÇAO DE NEGÓCIO/EXCLUÍDO TRANSAÇÃO/ ALTERADA EM : " & now & " / ALTERADA POR :" & request("chave")
	mensagem = "A função " & funcao & " utilizada para contrução do macro perfil " & macro & " foi alterada com a exclusão da transação " & transacao & " ."	
	correio.from = origem
	correio.AddAddress destino
	correio.Subject = assunto
	correio.body = mensagem		
	if destino<>"" then
		correio.send
	end if	
	set correio=nothing
end sub

rsdMacroPerfil.close
set rsdMacroPerfil = Nothing
conn_db.close
set conn_db = Nothing
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
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("txtAcao") <> "0" then
   str_Acao = request("txtAcao")
else
   str_Acao = ""
end if
if request("txtMacroPerfil") <> 0 then
   str_MacroPerfil = request("txtMacroPerfil")
else
   str_MacroPerfil = ""
end if
if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = ""
end if
if request("selSubModulo") <> 0 then
   str_SubModulo = request("selSubModulo")
else
   str_SubModulo = ""
end if

if request("txtPrefixoNomeTecnico") <> "0" then
   str_PrefixoNomeTecnico = UCase(Trim(request("txtPrefixoNomeTecnico")))
else
   str_PrefixoNomeTecnico = ""
end if
if request("txtNomeTecnico") <> "0" then
   str_NomeTecnico = UCase(Trim(request("txtNomeTecnico")))
else
   str_NomeTecnico = ""
end if
if request("txtDescMacroPerfil") <> "0" then
   str_DescMacroPerfil = Ucase(Trim(request("txtDescMacroPerfil")))
else
   str_DescMacroPerfil = ""
end if

if request("txtDescDetalhada") <> "0" then
   str_DescDetaMacroPerfil = Ucase(Trim(request("txtDescDetalhada")))
else
   str_DescDetaMacroPerfil = ""
end if
if request("txtEspecificacao") <> "0" then
   str_Especificacao = Ucase(Trim(request("txtEspecificacao")))
else
   str_Especificacao = ""
end if
if request("selFuncPrinc") <> "0" then
   str_FuncPrinc = Ucase(Trim(request("selFuncPrinc")))
else
   str_FuncPrinc = ""
end if
if request("txtSituacao") <> "" then
   str_SituacaoAtu = Ucase(Trim(request("txtSituacao")))
else
   str_SituacaoAtu = ""
end if
if request("txtNomeTecnico_Original") <> "" then
   str_NomeTecnico_Original = Ucase(Trim(request("txtNomeTecnico_Original")))
else
   str_NomeTecnico_Original = ""
end if
if request("txtDescMacroPerfil_Original") <> "" then
   str_DescMacroPerfil_Original = Ucase(Trim(request("txtDescMacroPerfil_Original")))
else
   str_DescMacroPerfil_Original = ""
end if
if request("txtDescDetalhada_Original") <> "" then
   str_DescDetaMacroPerfil_Original = Ucase(Trim(request("txtDescDetalhada_Original")))
else
   str_DescDetaMacroPerfil_Original = ""
end if
if request("txtEspecificacao_Original") <> "" then
   str_Especificacao_Original = Ucase(Trim(request("txtEspecificacao_Original")))
else
   str_Especificacao_Original = ""
end if
'========================= TESTA EXISTENCIA DE MACRO ==========================
str_SQL = ""
str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL "
str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
str_SQL = str_SQL & " WHERE MCPE_TX_NOME_TECNICO = '" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
if str_Acao <> "C" then 
   str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
end if
set rds_Existe_Funcao = db.Execute(str_SQL)
str_Nao_Existe = 1 ' indica repetição
if rds_Existe_Funcao.EOF then 
   str_Nao_Existe = 0 ' sem proplema para criação do MACRO
else 
   str_SQL = ""
   str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL "
   str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
   str_SQL = str_SQL & " WHERE MCPE_TX_NOME_TECNICO = '" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
   str_SQL = str_SQL & " and (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO IN ('EX', 'MR', 'EL', 'EP', 'ER')) "
   if str_Acao <> "C" then 
      str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
   end if     
   set rds_Existe_Funcao2 = db.Execute(str_SQL)
   if not rds_Existe_Funcao2.EOF then			      
      str_SQL = ""
	  str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL, MCPE_TX_DESC_MACRO_PERFIL "
      str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
      str_SQL = str_SQL & " WHERE MCPE_TX_NOME_TECNICO = '" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
	  str_SQL = str_SQL & " and (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO not IN ('EX', 'MR', 'EL', 'EP', 'ER')) "
      if str_Acao <> "C" then 
         str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
      end if	  
      set rds_Existe_Funcao2 = db.Execute(str_SQL)
	  if rds_Existe_Funcao2.EOF then
         str_Nao_Existe = 0 ' sem proplema para criação do MACRO
	  else
	     str_Desc_Macro_Perfil = rds_Existe_Funcao2("MCPE_TX_DESC_MACRO_PERFIL")
      end if 
   end if
end if   	  

if str_Nao_Existe = 1 then
   response.redirect "msg_ja_existe.asp?opt=0&txtTitFuncao=" & str_Desc_Macro_Perfil
end if
'========================= FIM TESTA EXISTENCIA DE MACRO ==========================
'========================= TESTA EXISTENCIA DE MACRO x FUNCAO ==========================
str_SQL = ""
str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL "
str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "'"
if str_Acao <> "C" then 
   str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
end if
set rds_Existe_Funcao = db.Execute(str_SQL)
str_Nao_Existe = 1 ' indica repetição
if rds_Existe_Funcao.EOF then 
   str_Nao_Existe = 0 ' sem proplema para criação do MACRO
else 
   str_SQL = ""
   str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL "
   str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "'"
   str_SQL = str_SQL & " and (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO IN ('EX', 'MR', 'EL', 'EP', 'ER')) "
   if str_Acao <> "C" then 
      str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
   end if   
   set rds_Existe_Funcao2 = db.Execute(str_SQL)
   if not rds_Existe_Funcao2.EOF then			      
      str_SQL = ""
	  str_SQL = str_SQL & " SELECT distinct MCPR_NR_SEQ_MACRO_PERFIL, MCPE_TX_DESC_MACRO_PERFIL "
      str_SQL = str_SQL & " FROM dbo.MACRO_PERFIL "
      str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "'"
	  str_SQL = str_SQL & " and (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO not IN ('EX', 'MR', 'EL', 'EP', 'ER')) "
      if str_Acao <> "C" then 
         str_SQL = str_SQL & " and MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
      end if	  
      set rds_Existe_Funcao2 = db.Execute(str_SQL)
	  if rds_Existe_Funcao2.EOF then
         str_Nao_Existe = 0 ' sem proplema para criação do MACRO
	  else
	     str_Desc_Macro_Perfil = rds_Existe_Funcao2("MCPE_TX_DESC_MACRO_PERFIL")
      end if 
   end if
end if   	  

if str_Nao_Existe = 1 then
   response.redirect "msg_ja_existe.asp?opt=4&txtTitFuncao=" & str_Desc_Macro_Perfil
end if
'========================= FIM TESTA EXISTENCIA DE MACRO X FUNCAO ==========================
' "C" opção de CRIACAO =======================
if str_Acao = "C" then 
   str_Desc_Acao = "Criação "
   set temp=db.execute("SELECT MAX(MCPR_NR_SEQ_MACRO_PERFIL)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_PERFIL")
   if not isnull(temp("codigo")) then
      sequencia=temp("CODIGO")+1
   else
      sequencia=1
   end if
   temp.close
   set temp = Nothing
   str_data = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   ssql=""
   ssql=ssql+ "INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL  "
   ssql=ssql+ "(MCPR_NR_SEQ_MACRO_PERFIL"
   ssql=ssql+ " , MCPE_TX_DESC_MACRO_PERFIL"
   ssql=ssql+ " , MCPE_TX_DESC_DETA_MACRO_PERFIL"      
   ssql=ssql+ " , MCPE_TX_NOME_TECNICO"
   ssql=ssql+ " , MCPE_TX_ESPECIFICACAO"   
   ssql=ssql+ " , MCPE_TX_SITUACAO"
   ssql=ssql+ " , MEPR_CD_MEGA_PROCESSO"
   ssql=ssql+ " , SUMO_NR_CD_SEQUENCIA"   
   ssql=ssql+ " , FUNE_CD_FUNCAO_NEGOCIO " 
   ssql=ssql+ " , ATUA_TX_OPERACAO"
   ssql=ssql+ " , ATUA_CD_NR_USUARIO "
   ssql=ssql+ " , ATUA_DT_ATUALIZACAO )"
   ssql=ssql+ " VALUES( " 
   ssql=ssql+ "" & sequencia & ""
   ssql=ssql+ ",'" & ucase(str_DescMacroPerfil) & "'"
   ssql=ssql+ ",'" & ucase(str_DescDetaMacroPerfil) & "'"
   ssql=ssql+ ",'" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
   ssql=ssql+ ",'" & ucase(str_Especificacao) & " - " & Session("CdUsuario") & " - " & str_data & "'"   
   ssql=ssql+ ",'EE'"
   ssql=ssql+ "," & str_MegaProcesso 
IF    str_SubModulo <> "" THEN
   ssql=ssql+ "," & str_SubModulo  
ELSE
   ssql=ssql+ ",NULL "  
END IF        
   ssql=ssql+ ",'" & str_FuncPrinc & "'"
   ssql=ssql+ ",'I','" & Session("CdUsuario") & "',GETDATE())"
   'response.Write(ssql)
   str_Status_1 = "EE"
   db.execute(ssql)   
   ' ================================== HERDA AS TRANSAÇÕES DA FUNÇÃO =================
	' ALTERADO PARA PEGAR APENAS AS TRANSAÇÕES INDEPENDENTE DO MEGA PROCESSO
	'==================================================================================	
   'str_SQL = ""
   'str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO( "
   'str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO "
   'str_SQL = str_SQL & " , MCPT_NR_SITUACAO_ALTERACAO, MCPT_NR_SITUACAO_ALTERACAO1, MCPT_NR_SITUACAO_ALTERACAO_FUNC "
   'str_SQL = str_SQL & " , ATUA_TX_OPERACAO "
   'str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
   'str_SQL = str_SQL & " (SELECT DISTINCT " & sequencia & ", TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO, "
   'str_SQL = str_SQL & " 0, 0, 0, "
   'str_SQL = str_SQL & " 'I', '" & Session("CdUsuario") &  "', GETDATE() "
   'str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   'str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "') "
   'db.execute(str_SQL)

   str_SQL = ""
   str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO( "
   str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO "
   str_SQL = str_SQL & " , MCPT_NR_SITUACAO_ALTERACAO, MCPT_NR_SITUACAO_ALTERACAO1, MCPT_NR_SITUACAO_ALTERACAO_FUNC "
   str_SQL = str_SQL & " , MCPT_NR_SITUACAO_PROCESSAMENTO "
   str_SQL = str_SQL & " , ATUA_TX_OPERACAO "
   str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
   str_SQL = str_SQL & " (SELECT DISTINCT " & sequencia & ", TRAN_CD_TRANSACAO, "
   str_SQL = str_SQL & " 0, 0, 0, 0, "
   str_SQL = str_SQL & " 'I', '" & Session("CdUsuario") &  "', GETDATE() "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "') "
   db.execute(str_SQL)

   ' ================================== VERIFICA OS DONOS =================
   str_SQL = ""
   str_SQL = str_SQL & " Select distinct TRAN_CD_TRANSACAO "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & sequencia
   'RESPONSE.WRITE "  SQL 1   "
   'RESPONSE.WRITE STR_SQL   
   'set rs_QtdMega=db.execute(str_SQL)
   int_Qtd_Mega = 0
   str_ListaMega = ""
   str_Necessita_Autorizacao = 0
   '
   ' rotina retiradA conforme autorizado pelo MARCELO / KATIA
   '
   'Do while not rs_QtdMega.EOF	   
   '   str_SQL = ""
	'  str_SQL = str_SQL & " SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
	'  str_SQL = str_SQL & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
    '  str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_MEGA INNER JOIN"
    '  str_SQL = str_SQL & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
    '  str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
    '  str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO_MEGA.TRAN_CD_TRANSACAO = '" & rs_QtdMega("TRAN_CD_TRANSACAO") & "'" 
	'  str_SQL = str_SQL & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
  
      'RESPONSE.WRITE "  SQL 2   "
      'RESPONSE.WRITE STR_SQL
				   
	'  Set rdsExiste2 = db.Execute(str_SQL)				   
	'  loo_Existe = False
	'  IF not rdsExiste2.EOF then
	'     Do While not rdsExiste2.EOF
	'        'if Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) = Trim(str_MegaProcesso) then
	'         if InStr("," & Session("AcessoUsuario") & ",","," &  Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) & ",") <> 0 then						 
	'           loo_Existe = True
    '           exit do
	'        end if
	'		rdsExiste2.Movenext
	'	 Loop
	'  else
	'     loo_Existe = True	 
	'  end if
	'  if loo_Existe = False then 'and not rdsExiste2.eof
    '     rdsExiste2.MoveFirst
    '     if not rdsExiste2.EOF then
	'        Do While not rdsExiste2.EOF
    '           Call Grava_Para_Autorizar(sequencia, rs_QtdMega("TRAN_CD_TRANSACAO"), rdsExiste2("MEPR_CD_MEGA_PROCESSO"))
	'		   str_Necessita_Autorizacao = 1
	'           rdsExiste2.Movenext
	'	    Loop
	'      end if		
	'  end if
	'  rdsExiste2.close
	'  set rdsExiste2 = Nothing
	'  rs_QtdMega.movenext	   
  'loop
 
 '  if str_Necessita_Autorizacao = "0" then 
 '     str_SQL = ""
'	  str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET " 
'	  str_SQL = str_SQL & " MCPE_TX_SITUACAO = 'EC' "
'	  str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & sequencia
'	  str_Status_1 = "EC"
 '     db.execute(str_SQL)
  ' end if	  
  
   ' ================================== GRAVA OS OBJETOS DAS TRANSAÇOES =================
'   str_SQL = ""
'   str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO( "
'   str_SQL = str_SQL & " MPTO_TX_SIT_ALTERACAO_VALOR, MPTO_TX_SIT_ALTERACAO_VALOR1, MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, "
'   str_SQL = str_SQL & " TROB_TX_CAMPO, TROB_TX_OBJETO, MPTO_TX_VALORES, TROB_TX_CRITICO , ATUA_TX_OPERACAO, "
'   str_SQL = str_SQL & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
'   str_SQL = str_SQL & " (SELECT '0','0', " & sequencia & ", TRAN_CD_TRANSACAO, TROB_TX_CAMPO, "
'   str_SQL = str_SQL & " TROB_TX_OBJETO, TRON_TX_VALORES, TROB_TX_CRITICO, "
'   str_SQL = str_SQL & " 'I', '" & Session("CdUsuario") &  "', GETDATE() "
'   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_OBJETO "
'   str_SQL = str_SQL & " WHERE TRAN_CD_TRANSACAO IN (SELECT TRAN_CD_TRANSACAO "
'   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
'   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "')) "

   'RESPONSE.WRITE " ***   TRANSACAO_OBJETO ****   "
   'RESPONSE.WRITE str_SQL

'   db.execute(str_SQL)
   
   COMENTARIO = "NOVO MACRO PERFIL"
   SSQL=""
   SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO("
   SSQL=SSQL+"MEPR_CD_MEGA_PROCESSO"
   SSQL=SSQL+", MHVA_TX_COMENTARIO"
   SSQL=SSQL+", MHVA_NR_SEQUENCIA_HIST"
   SSQL=SSQL+", MCPR_NR_SEQ_MACRO_PERFIL"
   SSQL=SSQL+", MHVA_TX_SITUACAO_MACRO"
   SSQL=SSQL+", ATUA_TX_OPERACAO"
   SSQL=SSQL+", ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO "
   SSQL=SSQL+" ) VALUES ( " & str_MegaProcesso & ",'" & COMENTARIO & "', " & "1" & ", " & sequencia & ", '"& str_Status_1 & "', 'I', '" & Session("CdUsuario") & "', GETDATE())"        		
   db.execute(ssql)
   'RESPONSE.Write(str_MegaProcesso)
'**********************************************************************************************
' caso seja 	MACRO PERFIL DE BW - GRAVA AUTOMATICAMENTE O MICRO
'**********************************************************************************************
   if str_MegaProcesso = 15 then
      set temp=db.execute("SELECT MAX(MIPE_NR_SEQ_MICRO_PERFIL)AS CODIGO FROM " & Session("PREFIXO") & "MICRO_PERFIL_R3")
      if not isnull(temp("codigo")) then
         sequencia2=temp("CODIGO")+1
      else
         sequencia2=1
      end if
      temp.close
      set temp = Nothing
   
      str_SQL = ""
	  str_SQL = str_SQL & " INSERT INTO MICRO_PERFIL_R3 ("
      str_SQL = str_SQL & " MIPE_NR_SEQ_MICRO_PERFIL"
	  str_SQL = str_SQL & " , MIPE_TX_NOME_TECNICO"
	  str_SQL = str_SQL & " , MIPE_TX_DESC_MICRO_PERFIL"
	  str_SQL = str_SQL & " , MIPE_TX_DESC_DETALHADA"
	  str_SQL = str_SQL & " , MCPR_NR_SEQ_MACRO_PERFIL"
	  str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
	  str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
	  str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
	  str_SQL = str_SQL & " ) VALUES ( "
	  str_SQL = str_SQL & sequencia2
      str_SQL = str_SQL & ",'" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
      str_SQL = str_SQL & ",'" & ucase(str_DescMacroPerfil) & "'"
      str_SQL = str_SQL & ",'" & ucase(str_DescDetaMacroPerfil) & "'"
	  str_SQL = str_SQL & "," & sequencia
      str_SQL = str_SQL & ",'I','" & Session("CdUsuario") & "',GETDATE())"
	  'response.Write(str_SQL)
      db.execute(str_SQL)

   end if

   db.Close
   set db = Nothing

	'response.Write("selMegaProcesso=" & str_MegaProcesso & "&selFuncao=" & str_FuncPrinc & "&txtMacroPerfil=" & sequencia & "&txtNomeTecnico=" & str_PrefixoNomeTecnico & str_NomeTecnico & "&txtOPT=1")
   
   ' ============== CHAMA TELA DE SELEÇÃO DE TRANSAÇÃO PARA MANUTENÇÃO =================
  ' response.redirect "rel_funcao_transacao.asp?selMegaProcesso=" & str_MegaProcesso & "&selFuncao=" & str_FuncPrinc & "&txtMacroPerfil=" & sequencia & "&txtNomeTecnico=" & str_PrefixoNomeTecnico & str_NomeTecnico & "&txtOPT=1"
else
   str_Desc_Acao = "Alteração "
   str_Criado = 0
   str_Necessita_Autorizacao = 0
   If Len(str_Especificacao) <> 0 OR Trim(str_DescDetaMacroPerfil_Original) <> Trim(str_DescDetaMacroPerfil) OR Trim(str_DescMacroPerfil_Original) <> Trim(str_DescMacroPerfil) OR Trim(str_NomeTecnico_Original) <> Trim(str_NomeTecnico) then 
      If str_SituacaoAtu = "CR" or str_SituacaoAtu = "AP" or str_SituacaoAtu = "AR"  then
	     ' CRIADO - ALTERADO NO R3 - EM ALTERAÇÃO NO R3
         str_SituacaoAtu = "AR"
		 str_Criado = 1
      end if
	  str_SQL = ""
	  str_SQl = str_SQL & " Select "
	  str_SQl = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, MAOA_TX_AUTORIZADO "
	  str_SQl = str_SQL & " FROM dbo.MACRO_OBJ_AUTORIZA "
      str_SQl = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
	  set rdsAutorizacao = db.execute(str_SQL)
	  if not rdsAutorizacao.EOF then
	     str_Necessita_Autorizacao = 1
	     str_SQL = ""
	     str_SQl = str_SQL & " Select "
	     str_SQl = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, MAOA_TX_AUTORIZADO "   
	     str_SQl = str_SQL & " FROM dbo.MACRO_OBJ_AUTORIZA "
         str_SQl = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
		 str_SQl = str_SQL & " and MAOA_TX_AUTORIZADO in (0,2)"
	     set rdsAutorizacaoDada = db.execute(str_SQL)
	     if not rdsAutorizacaoDada.EOF then
		    str_Necessita_Autorizacao = 2
		 else
		    str_Necessita_Autorizacao = 0
		 end if
	  else
	     str_Necessita_Autorizacao = 0
	  end if
      rdsAutorizacao.close
	  set rdsAutorizacao = Nothing
	  'response.Write(Trim(ucase(str_PrefixoNomeTecnico)) & Trim(ucase(str_NomeTecnico)))
	  if Len(str_Especificacao) <> 0 then
	     str_data = Day(date()) & "/" & Month(date()) & "/" & Year(date())
	     str_espec = ucase(str_Especificacao_Original) & "    **/**   " & ucase(str_Especificacao) & " - " & Session("CdUsuario") & " - " & str_data 
      else
	  	 str_espec = ucase(str_Especificacao_Original)
	  end if
      ssql=""
      ssql=ssql+ " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET  "
      ssql=ssql+ " MCPE_TX_DESC_MACRO_PERFIL = '" & ucase(str_DescMacroPerfil) & "'" 
      ssql=ssql+ " , MCPE_TX_DESC_DETA_MACRO_PERFIL = '" & ucase(str_DescDetaMacroPerfil) & "'" 	  	  
      ssql=ssql+ " , MCPE_TX_NOME_TECNICO = '" & Trim(ucase(str_PrefixoNomeTecnico)) & Trim(ucase(str_NomeTecnico)) & "'"
      ssql=ssql+ " , MCPE_TX_ESPECIFICACAO = '" & str_espec & "'" 	  	  
      ssql=ssql+ " , MCPE_TX_SITUACAO = '" &  str_SituacaoAtu & "'"
      ssql=ssql+ " , ATUA_TX_OPERACAO = 'A'"
      ssql=ssql+ " , ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
      ssql=ssql+ " , ATUA_DT_ATUALIZACAO = GETDATE()"
      ssql=ssql+ " where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
	  'response.Write(ssql)	  
      db.execute(ssql)	  
      sequencia = str_MacroPerfil
	  'response.Write("<p>" & " original " & Trim(str_DescMacroPerfil_Original))
	  'response.Write("<p>" & " novo " & Trim(str_DescMacroPerfil))
	  COMENTARIO = " NÃO IDENTIFICADO A ALTERAÇÃO"
      if Trim(str_DescDetaMacroPerfil_Original) <> Trim(str_DescDetaMacroPerfil) then 
         COMENTARIO = "ALTERADO A DESCRIÇÃO DETALHADA DO MACRO PERFIL"
	  elseif Trim(str_NomeTecnico_Original) <> Trim(str_NomeTecnico) then 
         COMENTARIO = "ALTERADO O NOME TECNICO DO MACRO PERFIL"
      elseIf Trim(str_DescMacroPerfil_Original) <> Trim(str_DescMacroPerfil) then
         COMENTARIO = "ALTERADO DESCRIÇÃO DO MACRO PERFIL"	  
	  end if

      SSQL=""
      SSQL=SSQL + " SELECT Max(MHVA_NR_SEQUENCIA_HIST) as MAX_SEQ FROM MACRO_HISTORICO_VALIDACAO "
	  SSQL=SSQL + " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & sequencia	  
	  'response.Write("<p>" & " aaaaa  =" & ssql)
	  	  
	  Set Num_Seq = db.Execute(SSQL)
      'if Num_Seq.EOF then
	  if IsNull(Num_Seq("MAX_SEQ")) then
	     int_MaxSeq = 1	
	  else
	     int_MaxSeq = Num_Seq("MAX_SEQ") + 1	
	  end if
	  'response.Write("<p>" & " ccccc  =" & Num_Seq("MAX_SEQ"))
	  Num_Seq.Close
	  set Num_Seq = Nothing
	  'response.Write("<p>" & " bbbbb  =" & int_MaxSeq)
      SSQL=""
      SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO("
      SSQL=SSQL+" MEPR_CD_MEGA_PROCESSO"
      SSQL=SSQL+", MHVA_TX_COMENTARIO"
      SSQL=SSQL+", MHVA_NR_SEQUENCIA_HIST"
      SSQL=SSQL+", MCPR_NR_SEQ_MACRO_PERFIL"
      SSQL=SSQL+", MHVA_TX_SITUACAO_MACRO"
      SSQL=SSQL+", ATUA_TX_OPERACAO"
      SSQL=SSQL+", ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO "
      SSQL=SSQL+" ) VALUES ( " & str_MegaProcesso & ",'" & COMENTARIO & "', " & int_MaxSeq & ", " & sequencia & ", '"& str_SituacaoAtu & "', 'I', '" & Session("CdUsuario") & "', GETDATE())"        		
	  'response.Write(SSQL)
      db.execute(ssql)
   end if	  
   db.Close
   set db = Nothing
  ' response.redirect "msg_ja_existe.asp?opt=1"  
  sequencia = str_MacroPerfil 
end if


SUB GRAVA_LOG(COD,TABELA,OPE,CONECTA)
    TX_CHAVE_TABELA=COD
    TX_TABELA=TABELA
    TX_OPERACAO=OPE
    SSQL_LOG=""
    SSQL_LOG="INSERT INTO " & Session("PREFIXO") & "LOG_GERAL(LOGE_TX_CHAVE_TABELA,LOGE_TX_TABELA,LOGE_DT_DATA_LOG,LOGE_TX_OPERACAO,LOGE_TX_USUARIO) "
    SSQL_LOG=SSQL_LOG+"VALUES('" & TX_CHAVE_TABELA & "', "
    SSQL_LOG=SSQL_LOG+"'" & TX_TABELA & "', "
    SSQL_LOG=SSQL_LOG+"GETDATE(), "
    SSQL_LOG=SSQL_LOG+"'" & TX_OPERACAO & "', "
    SSQL_LOG=SSQL_LOG+"'" & Session("CdUsuario") & "')"
    IF CONECTA=1 THEN
       DB.EXECUTE(SSQL_LOG)
    ELSE
	   CONN_DB.EXECUTE(SSQL_LOG)
    END IF
END SUB

'call grava_log(str_FuncPrinc,"" & Session("PREFIXO") & "MACRO_PERFIL","I",1)

Sub Grava_Para_Autorizar(pMacro, pTransacao, pMega)
	str_SQL = ""
	str_SQL = str_SQL & " Insert into " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA( "
    str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL "
	str_SQL = str_SQL & " , TRAN_CD_TRANSACAO"
	str_SQL = str_SQL & " , MEPR_CD_MEGA_PROCESSO"
	str_SQL = str_SQL & " , MAOA_TX_AUTORIZADO"
	str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
	str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
	str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO "
	str_SQL = str_SQL & " ) values ( "
	str_SQL = str_SQL & pMacro 
	str_SQL = str_SQL & ",'" & pTransacao & "'"
	str_SQL = str_SQL & "," & pMega
	str_SQL = str_SQL & ",'0'" 
    str_SQL = str_SQL & ",'I'"
    str_SQL = str_SQL & ",'" & Session("CdUsuario") & "',GETDATE())"
	'RESPONSE.WRITE STR_SQL
	db.execute(str_SQL)
end sub
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="js/troca_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="../Curso/valida_cad_curso.asp" name="frm1">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      <td colspan="3" height="20">&nbsp;</td>
  </tr>
</table>
        
  <table width="847" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <div align="center"><font color="#330099" size="3" face="Verdana">Grava&ccedil;&atilde;o 
          de Macro Perfil</font></div>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table width="78%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"><%'=str_Acao%> </td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"><%'=str_Necessita_Autorizacao%></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"><%'=str_Criado%></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><%=str_Desc_Acao%> de Macro Perfil com sucesso!</b></font> </td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"> <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
	      <% if str_Acao = "C" then %>
          <tr> 
            <td height="41"><a href="incluir_macro_perfil.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de cria&ccedil;&atilde;o de Macro Perfil</font></td>
          </tr>		  
  	      <% end if
		  if str_Acao = "M" then %>
          <tr> 
            <td height="41"><a href="seleciona_macro_perfil.asp?pOPT=1"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de altera&ccedil;&atilde;o de Macro Perfil</font></td>
          </tr>
          <% end if
		  if (str_Acao = "C" and str_Necessita_Autorizacao <> "0") or (str_Acao = "M" and str_Necessita_Autorizacao <> "0") then %>
          <tr> 
            <td height="41"><a href="valida_status.asp?txtOrigem=0&opt=EA&amp;macro=<%=sequencia%>&amp;acao=<%=str_acao%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Envia 
              para valida&ccedil;&atilde;o</font></td>
          </tr>
          <% elseif (str_Acao = "M" and str_Criado = 1 ) then 
		            str_Msg = "Macro Pefil já criado no R/3. Foi encaminhado automaticamente para alteração no R/3."
		  %>
          <% elseif ((str_Acao = "C" and str_Necessita_Autorizacao = "0") or (str_Acao = "M" and str_Necessita_Autorizacao = "0")) and str_Criado = 0  then %>
          <tr> 
            <td height="41"><a href="valida_status.asp?txtOrigem=0&opt=EC&amp;macro=<%=sequencia%>&amp;acao=<%=str_acao%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font color="#003366" size="2" face="Verdana, Arial, Helvetica, sans-serif">Envia 
              para cria&ccedil;&atilde;o no R/3</font></td>
          </tr>
          <% end if %>
          <tr> 
            <td height="41"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela Principal</font></td>
          </tr>
          <tr> 
            <td height="41">&nbsp;</td>
            <td height="41"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=str_Msg%> </strong></font></td>
          </tr>
        </table></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="70%"> </td>
      <td width="16%">&nbsp;</td>
    </tr>
  </table>
</form>

</body>

</html>

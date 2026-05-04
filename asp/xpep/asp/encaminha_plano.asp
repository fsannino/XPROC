<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_Cd_Onda = request("selOnda")
str_Fase = request("selFases")

Dim strValorSelTask1, str_Atividade, str_Reserved_Data

if str_Cd_Onda = "7" and Request("selTaskSub") <> "" then
	strValorSelTask1 = split(Request("selTaskSub"), "|")
	str_Atividade = strValorSelTask1(0)
	str_Reserved_Data = strValorSelTask1(1)
		
	strIdAtividade = split(Request("selTask1"), "|")
	intIdAtividade = strIdAtividade(0)		
	
	Response.write "1 " & str_Atividade & "<br>"
	
else
	if Request("selTask1") <> "" then
		strValorSelTask1 = split(Request("selTask1"), "|")
		str_Atividade = strValorSelTask1(0)
		str_Reserved_Data = strValorSelTask1(1)
	else
		str_Atividade = ""
		str_Reserved_Data = ""
	end if
	Response.write "2 " & str_Atividade & "<br>"
end if

strPlanoOriginal = Request("pPlanoOriginal")

if Request("selPlano") <> "" then
	strValorPlano =  split(Request("selPlano"), "|")
	str_Cd_Plano = strValorPlano(0)
end if

if Request("selPlano2") <> "" then
	strValorPlano2 =  split(Request("selPlano2"), "|")
	str_Cd_Plano2 = strValorPlano2(0)
end if
	
'Response.write str_Cd_Plano & "<br>"
'Response.write str_Cd_Plano2 & "<br>"
'Response.end

str_SiglaPlano2 = request("pSiglaPlano")
str_TipoCadastramento = request("selTipoCadastro")
str_Plano_Origem = request("pPlano_Origen")

str_TpPlano = ""
str_TpPlano = str_TpPlano & "Select PLAN_TX_SIGLA_PLANO, PLAN_NR_CD_PROJETO_PROJECT "
str_TpPlano = str_TpPlano & " From XPEP_PLANO_ENT_PRODUCAO "
str_TpPlano = str_TpPlano & " WHERE "
str_TpPlano = str_TpPlano & " PLAN_NR_SEQUENCIA_PLANO = " & Trim(str_Cd_Plano)
'RESPONSE.Write(str_TpPlano)

set rdsTpPlano = db_Cogest.Execute(str_TpPlano)
if not rdsTpPlano.Eof then
   str_SiglaPlano = rdsTpPlano("PLAN_TX_SIGLA_PLANO")
   int_Cd_Projeto_Project = rdsTpPlano("PLAN_NR_CD_PROJETO_PROJECT")   
else
   str_SiglaPlano = ""
   int_Cd_Projeto_Project = ""
end if
'response.Write "int_Cd_Projeto_Project - " & int_Cd_Projeto_Project & "<br>"
'RESPONSE.End()
'response.Write " str_TpPlano - " & str_TpPlano & "<br>"
'response.Write " str_Plano_Origem - " & str_Plano_Origem & "<br>"
'response.Write "str_SiglaPlano - " & str_SiglaPlano & "<br>"
'response.Write "str_Cd_Plano - " & str_Cd_Plano & "<br>"
'response.Write "str_Atividade - " & str_Atividade & "<br>"
'response.Write "int_Cd_Projeto_Project - " & int_Cd_Projeto_Project & "<br>"
'Response.write "str_TipoCadastramento - " & str_TipoCadastramento & "<br>"
'Response.write "str_TipoCadastramento - " & str_TipoCadastramento & "<br>"
'Response.end

str_ExisteTarefaGeral = ""
str_ExisteTarefaGeral = str_ExisteTarefaGeral & " Select PLAN_NR_SEQUENCIA_PLANO "
str_ExisteTarefaGeral = str_ExisteTarefaGeral & " From XPEP_PLANO_TAREFA_GERAL " 
str_ExisteTarefaGeral = str_ExisteTarefaGeral & " where PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano	

'Response.write str_Atividade & "<br>"
'Response.write intIdAtividade & "<br>" 
'Response.end

if str_Atividade <> "" then
	str_ExisteTarefaGeral = str_ExisteTarefaGeral & " and PLTA_NR_SEQUENCIA_TAREFA = " & str_Atividade
else
	str_ExisteTarefaGeral = str_ExisteTarefaGeral & " and PLTA_NR_SEQUENCIA_TAREFA = " & intIdAtividade
end if

	
'if str_TipoCadastramento = "PCM" then	
	'set rdsExisteTarefaGeral = db_Cogest.Execute(str_ExisteTarefaGeral)

	'if rdsExisteTarefaGeral.eof then			
		'strMsg  = "Não existe detalhamento cadastrado para esta atividade!"	
		'response.redirect "msg_geral.asp?pMsg=" & strMsg
	'else	
		'str_ExisteTarefa = ""
		'str_ExisteTarefa = str_ExisteTarefa & " Select PLAN_NR_SEQUENCIA_PLANO "
		'str_ExisteTarefa = str_ExisteTarefa & " From XPEP_PLANO_TAREFA_PCM "
		'str_ExisteTarefa = str_ExisteTarefa & " where PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano
		'str_ExisteTarefa = str_ExisteTarefa & " AND PPCM_NR_SEQUENCIA_TAREFA = " & str_Atividade
	'end if
	'rdsExisteTarefaGeral.close
	'set rdsExisteTarefaGeral = nothing
'elseif str_TipoCadastramento = "PAC" then	
if str_TipoCadastramento = "PAC" then	
	'Response.write str_ExisteTarefaGeral
	'Response.end
	
	set rdsExisteTarefaGeral = db_Cogest.Execute(str_ExisteTarefaGeral)

	if rdsExisteTarefaGeral.eof then			
		strMsg  = "Não existe detalhamento cadastrado para esta atividade!"	
		response.redirect "msg_geral.asp?pMsg=" & strMsg
	else	
		str_ExisteTarefa = ""
		str_ExisteTarefa = str_ExisteTarefa & " Select PLAN_NR_SEQUENCIA_PLANO "
		str_ExisteTarefa = str_ExisteTarefa & " From XPEP_PLANO_TAREFA_PAC "
		str_ExisteTarefa = str_ExisteTarefa & " where PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano
		str_ExisteTarefa = str_ExisteTarefa & " AND PLTA_NR_SEQUENCIA_TAREFA = " & str_Atividade
		
		set rdsExisteTarefa = db_Cogest.Execute(str_ExisteTarefa)
		if not rdsExisteTarefa.Eof then
		   str_Acao = "A" '*** Encontrou
		else
		   str_Acao = "I" '*** Não Encontrou
		end if
	end if
	rdsExisteTarefaGeral.close
	set rdsExisteTarefaGeral = nothing	
elseif str_TipoCadastramento = "DET" then
	
	'select case str_SiglaPlano
	'  case "PAI"
		'	strTabela = "XPEP_PLANO_TAREFA_PAI"
	  'case "PCD"
			'strTabela = "XPEP_PLANO_TAREFA_PCD"
	 ' case "PDS"
			'strTabela = "XPEP_PLANO_TAREFA_PDS"
	'  case "PPO"   		
		'	strTabela = "XPEP_PLANO_TAREFA_PPO"
	'end select
	
	'str_ExisteTarefa = ""
	'str_ExisteTarefa = str_ExisteTarefa & " Select GERAL.PLAN_NR_SEQUENCIA_PLANO "
	'str_ExisteTarefa = str_ExisteTarefa & " From XPEP_PLANO_TAREFA_GERAL GERAL, " & strTabela & " PLANO"
	'str_ExisteTarefa = str_ExisteTarefa & " where GERAL.PLTA_NR_SEQUENCIA_TAREFA = PLANO.PLTA_NR_SEQUENCIA_TAREFA " 
	'str_ExisteTarefa = str_ExisteTarefa & " and GERAL.PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano	
	'str_ExisteTarefa = str_ExisteTarefa & " and GERAL.PLTA_NR_SEQUENCIA_TAREFA = " & str_Atividade
	
	str_ExisteTarefa = ""
	str_ExisteTarefa = str_ExisteTarefa & " Select PLAN_NR_SEQUENCIA_PLANO "
	str_ExisteTarefa = str_ExisteTarefa & " From XPEP_PLANO_TAREFA_GERAL " 
	str_ExisteTarefa = str_ExisteTarefa & " where PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano	
	str_ExisteTarefa = str_ExisteTarefa & " and PLTA_NR_SEQUENCIA_TAREFA = " & str_Atividade
	
	set rdsExisteTarefa = db_Cogest.Execute(str_ExisteTarefa)
	if not rdsExisteTarefa.Eof then
	   str_Acao = "A" '*** Encontrou
	else
	   str_Acao = "I" '*** Não Encontrou
	end if
end if
	
	'Response.write str_ExisteTarefa
	'Response.end
	
if str_Plano_Origem = "" then
	str_Plano_Origem = str_SiglaPlano
end if

'Response.write "str_Acao -" & str_Acao & "<br>"
'Response.write str_ExisteTarefa	
'Response.end

if str_SiglaPlano2 <> "" then
   str_SiglaPlano = str_SiglaPlano2
end if

if str_SiglaPlano = "PCM" then
	response.redirect "inclui_altera_plano_pcm.asp?pOnda=" & str_Cd_Onda & "&pPlano=" & str_Cd_Plano & "&pPlano2=" & str_Cd_Plano2 & "&pPlano_Origem=" & str_Plano_Origem & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal
end if

'if str_TipoCadastramento = "PCM" then
'	response.redirect "inclui_altera_plano_pcm.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pPlano_Origem=" & str_Plano_Origem
'elseif str_TipoCadastramento = "PAC" then	
if str_TipoCadastramento = "PAC" then	
	'response.redirect "inclui_altera_plano_pac.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pPlano_Origem=" & str_Plano_Origem & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal & "&pIdAtividade=" & intIdAtividade
	response.redirect "inclui_altera_plano_pac_1.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pPlano_Origem=" & str_Plano_Origem & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal & "&pIdAtividade=" & intIdAtividade
elseif str_TipoCadastramento = "DET" then
	if str_SiglaPlano <> "" then
	   Select case str_SiglaPlano
		   case "PAI"
				response.redirect "inclui_altera_plano_pai.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal
		   case "PCD"
		   		if str_Cd_Onda = "7" then
					response.redirect "inclui_altera_plano_pcd_rh.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal
		   		else
					response.redirect "inclui_altera_plano_pcd.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal
				end if
		   case "PDS"
				response.redirect "inclui_altera_plano_pds.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano	& "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal		   
		   case "PPO"		   		
				response.redirect "inclui_altera_plano_ppo.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal & "&pIdAtividade=" & intIdAtividade
		   'case "PCE"
		   		'response.redirect "inclui_altera_plano_pce.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade & "&pOnda=" & str_Cd_Onda & "&pResData=" & str_Reserved_Data & "&pPlano=" & str_Cd_Plano & "&pFase=" & str_Fase & "&pPlanoOriginal=" & strPlanoOriginal
		End Select    
	else
		response.redirect "msg_.asp?pAcao=" & str_Acao & "&pCdProjProject=" & int_Cd_Projeto_Project & "&pTArefa=" &  str_Atividade
	end if
end if
%>


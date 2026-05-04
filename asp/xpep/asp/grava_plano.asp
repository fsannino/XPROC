<%
Response.Expires=0

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim rdsMaxPlano, strPlano, intCDPlanoGeral

strGravado = 0

strAcao				= Trim(Request("pAcao"))
strPlano 			= Request("pPlano")
intPlano2 			= Request("pintPlano2")
intIdTaskProject	= Request("idTaskProject")
intPlano			= Request("pintPlano")
int_CD_Onda			= Request("pOnda")
str_Fase			= Request("pFase")

if Request("pNomeAtividade") <> "" then
	strNomeAtividade	= Request("pNomeAtividade")
	strDtInicioAtiv 	= Formatdatetime(Request("pDtInicioAtiv"), 2)
	strDtFimAtiv 		= Formatdatetime(Request("pDtFimAtiv"), 2)
end if

int_Cd_Projeto_Project 	= Request("pCdProjProject")
str_Cd_Onda 			= Request("pOnda")
str_Cd_Plano 			= Request("pPlano")
str_Fase 				= Request("pFase")
strNomeOnda 			= Request("pNomeOnda")
intOnda 				= Request("pOnda")

strNomePlanoOrigem		= Request("pPlano_Origem")

'Response.write "strNomePlanoOrigem -" & strNomePlanoOrigem & "<br>"
'Response.write "strPlano -" & strPlano & "<br>"
'Response.write "intPlano2 -" & intPlano2 & "<br>"
'Response.write "intIdTaskProject -" & intIdTaskProject & "<br>"
'response.write "intPlano -" & intPlano & "<br>"
'Response.write "strAcao - " & strAcao
'Response.write "strNomeOnda - " & strNomeOnda
'Response.write "intOnda - " & intOnda
'Response.end

strMSG =  ""
		
intCdSequencialPCM = Request("pCdPCM")
			
'************************************** INCLUSÃO ************************************************
if strAcao = "I" then

	blnNaoCadastraPlano = False
	
	if strPlano = "PCM" or strPlano = "PAC" then	
		'strVerificaAtvidade = ""
		'strVerificaAtvidade = strVerificaAtvidade & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
		'strVerificaAtvidade = strVerificaAtvidade & " FROM XPEP_PLANO_TAREFA_GERAL"
		'strVerificaAtvidade = strVerificaAtvidade & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & intIdTaskProject		
		'Set rdsVerificaAtvidade = db_Cogest.Execute(strVerificaAtvidade)	
		'if not rdsVerificaAtvidade.EOF then
			'blnNaoCadastraPlano = True
		'end if
		'rdsVerificaAtvidade.close
		'set rdsVerificaAtvidade = nothing
		
		strVerificaAtvidade = ""
		strVerificaAtvidade = strVerificaAtvidade & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
		strVerificaAtvidade = strVerificaAtvidade & " FROM XPEP_PLANO_TAREFA_GERAL"
		strVerificaAtvidade = strVerificaAtvidade & " WHERE PLTA_NR_ID_TAREFA_PROJECT = 999999999"
	else		
		strVerificaAtvidade = ""
		strVerificaAtvidade = strVerificaAtvidade & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
		strVerificaAtvidade = strVerificaAtvidade & " FROM XPEP_PLANO_TAREFA_GERAL"
		strVerificaAtvidade = strVerificaAtvidade & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & intIdTaskProject		
	end if	
	
	'Response.write strVerificaAtvidade
	'Response.end
	
	Set rdsVerificaAtvidade = db_Cogest.Execute(strVerificaAtvidade)			
	
	if not rdsVerificaAtvidade.EOF then
		strMSG = "Já existe detalhamento cadastrado para esta atividade."
	else	
		'*** Seleciona o cod para a Nova Tarefa Geral - na tabela XPEP_PLANO_TAREFA_GERAL		
		'intCDPlanoGeral = 0	
		'Set rdsMaxPlano = db_Cogest.Execute("SELECT MAX(PLTA_NR_SEQUENCIA_TAREFA) AS INT_MAIOR_TAREFA_GERAL FROM XPEP_PLANO_TAREFA_GERAL")			
		'if isnull(rdsMaxPlano("INT_MAIOR_TAREFA_GERAL")) then
		'	intCDPlanoGeral = 1
		'else
		'	intCDPlanoGeral = rdsMaxPlano("INT_MAIOR_TAREFA_GERAL") + 1
		'end if
		'rdsMaxPlano.Close	
		'set rdsMaxPlano = nothing
		
		'****** PEGARÁ O COD DA TAREFA LÁ DO PROJECT
		intCDPlanoGeral = intIdTaskProject
		
		'*** Query qu montra strSQL para inclusão do Plano
		strSQL_NovoPlanoGeral = ""
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " INSERT INTO XPEP_PLANO_TAREFA_GERAL( "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " PLAN_NR_SEQUENCIA_PLANO "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,PLTA_NR_SEQUENCIA_TAREFA "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,PLTA_NR_ID_TAREFA_PROJECT "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,PLTA_TX_DESC_ATIVIDADE "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,PLTA_DT_INICIO_ATIV "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,PLTA_DT_TERMINO_ATIV "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,ATUA_TX_OPERACAO "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,ATUA_CD_NR_USUARIO "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ,ATUA_DT_ATUALIZACAO "
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ) Values(" & intPlano & "," & intCDPlanoGeral & ","
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & intIdTaskProject & ",'" & strNomeAtividade & "',"
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'" & strDtInicioAtiv & "','" & strDtFimAtiv & "',"
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'I', '" & Session("CdUsuario") & "', GETDATE())" 
								
		'*** PLANO DE PARADA OPERACIONAL - INCLUSÃO
		if strPlano = "PPO" then
					
			txtDescrParada 		= Trim(Ucase(Request("txtDescrParada")))			
			strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
			strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
			strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))						
			intTempParada 		= Request("txtTempParada")
			strUnidadeMedida	= Request("selUnidMedida")			
			strProcedParada 	= Trim(Ucase(Request("txtProcedParada")))			
			strUsuarioGestor 	= Trim(Ucase(Request("txtUsuarioGestor")))			
			intIdAtividade		= Request("pIdAtividade")
										
			if intIdAtividade = "" then
				intIdAtividade = 0 							
			end if
																	
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtDataParada = 	split(Request("txtDtParadaLegado"),"/")	
			strDia = vetDtDataParada(0)
			strMes = vetDtDataParada(1)
			strAno = vetDtDataParada(2)	
			strDtParadaLegado = strMes & "/" & strDia & "/" & strAno 
					
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtDataR3 = 	split(Request("txtDtIniR3"),"/")	
			strDia = vetDtDataR3(0)
			strMes = vetDtDataR3(1)
			strAno = vetDtDataR3(2)	
			strDtIniR3 = strMes & "/" & strDia & "/" & strAno 
								
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtLimiteAprov = 	split(Request("txtDtLimiteAprov"),"/")	
			strDia = vetDtLimiteAprov(0)
			strMes = vetDtLimiteAprov(1)
			strAno = vetDtLimiteAprov(2)	
			strDtLimiteAprov = strMes & "/" & strDia & "/" & strAno 
			
			Call VerificaUsuarioExistente("Responsável Sinergia",strRespTecSinGeral, "Sinergia", "PPO")			
			Call VerificaUsuarioExistente("Responsável Legado - Técnico",strRespTecLegGeral, "Legado", "PPO")
			Call VerificaUsuarioExistente("Responsável Legado - Funcional",strRespFunLegGeral, "Legado", "PPO")
			Call VerificaUsuarioExistente("Gestor do Processo",strUsuarioGestor, "Legado", "PPO")
			
			'Response.write intIdTaskProject & "<br>"
			'Response.write intPlano & "<br>"
			'Response.write txtDescrParada & "<br>"
			'Response.write strRespTecSinGeral & "<br>"
			'Response.write strRespTecLegGeral & "<br>"
			'Response.write strRespFunLegGeral & "<br>"
			'Response.write intTempParada & "<br>"
			'Response.write strUnidadeMedida & "<br>"	
			'Response.write strProcedParada & "<br>"
			'Response.write strDtParadaLegado& "<br>"
			'Response.write strDtIniR3 & "<br>"
			'Response.write strUsuarioGestor & "<br>"
			'Response.write strDtLimiteAprov & "<br><br>"
			'Response.end	
												
			strSQL_NovoPlanoPPO = ""
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & " INSERT INTO XPEP_PLANO_TAREFA_PPO ( "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_NR_ID_ATIVIDADE "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_TX_DESCRICAO_PARADA "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_TX_QTD_TEMPO_PARADA "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_TX_UNID_TEMPO_PARADA "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_TX_PROCEDIMENTOS_PARADA "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_DT_PARADA_LEGADO "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_DT_INICIO_R3 "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_DT_LIMITE_APROVACAO "
			'strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_NR_ID_PLANO_CONTINGENCIA "
			'strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", PPPO_NR_ID_PLANO_COMUNICACAO "	
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", USUA_CD_USUARIO_RESP_SINER "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", USUA_CD_USUARIO_RESP_LEG_TEC "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", USUA_CD_USUARIO_RESP_LEG_FUN "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", USUA_CD_USUARIO_GESTOR_PROC "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & " ) Values(" & intPlano & ","
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & intCDPlanoGeral & "," & intIdAtividade & ",'" & txtDescrParada & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & intTempParada & ",'" & strUnidadeMedida & "','" & strProcedParada & "','" & strDtParadaLegado & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'" & strDtIniR3 & "','" & strDtLimiteAprov & "','" & strRespTecSinGeral & "','" & strRespTecLegGeral & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'" & strRespFunLegGeral & "','" & strUsuarioGestor & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
								
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPPO)
		
		'*** PLANO DE COMUNICAÇÃO - INCLUSÃO
		elseif strPlano = "PCM" then		
			
			strAtividade 		= Trim(Ucase(Request("txtAtividade")))
			strComunicacao		= Request("selComunicacao")			
			strOqueComunicar 	= Trim(Ucase(Request("txtOqueComunicar")))
			strAQuemComunicar 	= Trim(Ucase(Request("txtAQuemComunicar")))
			strUnidadeOrgao 	= Trim(Ucase(Request("txtUnidadeOrgao")))			
			strRespConteudo		= Trim(Ucase(Request("txtRespConteudo")))		
			strRespDivulg		= Trim(Ucase(Request("txtRespDivulg")))
			strComo				= Trim(Ucase(Request("txtComo")))		
			strAprovadorPB		= Trim(Ucase(Request("txtAprovadorPB")))
			strfilArquivo1		= Trim(Ucase(Request("filArquivo1")))		
			strfilArquivo2		= Trim(Ucase(Request("filArquivo2")))		
			strfilArquivo3		= Trim(Ucase(Request("filArquivo3")))		
											
			strDia = ""		
			strMes = ""
			strAno = ""			
			vetDtQuandoOcorre = split(Request("txtQuandoOcorre"),"/")	
			strDia = vetDtQuandoOcorre(0)
			strMes = vetDtQuandoOcorre(1)
			strAno = vetDtQuandoOcorre(2)	
			strQuandoOcorre = strMes & "/" & strDia & "/" & strAno 
														
			if trim(Request("txtDtAprovacao")) <> "" then																			
				strDia = ""		
				strMes = ""
				strAno = ""			
				vetDtAprovacao = 	split(Request("txtDtAprovacao"),"/")	
				strDia = vetDtAprovacao(0)
				strMes = vetDtAprovacao(1)
				strAno = vetDtAprovacao(2)	
				strDtAprovacao = strMes & "/" & strDia & "/" & strAno 			
			end if
			
			if strAprovadorPB <> "" then
				Call VerificaUsuarioExistente("Aprovador PB",strAprovadorPB, "Legado", "PCM")	
			end if
			
			'Response.write strComunicacao & "<br>"
			'Response.write strOqueComunicar & "<br>"
			'Response.write strAQuemComunicar & "<br>"
			'Response.write strUnidadeOrgao 	& "<br>"
			'Response.write strQuandoOcorre 	& "<br>"
			'Response.write strRespConteudo	& "<br>"
			'Response.write strRespDivulg	& "<br>"			
			'Response.write strDtAprovacao 	& "<br>"
			'Response.write strfilArquivo1	& "<br>"
			'Response.write strfilArquivo2	& "<br>"
			'Response.write strfilArquivo3	& "<br>"
			'Response.end
			
			sql_MaxPCM = "SELECT MAX(PPCM_NR_SEQUENCIA_TAREFA) AS NovoCdPCM FROM XPEP_PLANO_TAREFA_PCM"
			set rsdMaxPCM = db_Cogest.Execute(sql_MaxPCM)			
			if isnull(rsdMaxPCM("NovoCdPCM")) then
				intCdPCM = 1
			else
				intCdPCM = rsdMaxPCM("NovoCdPCM") + 1
			end if
			rsdMaxPCM.Close	
			set rsdMaxPCM = nothing
																				
			strSQL_NovoPlanoPCM = ""
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " INSERT INTO XPEP_PLANO_TAREFA_PCM ( "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_NR_SEQUENCIA_TAREFA "			
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_ATIVIDADE "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_TP_COMUNICACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_O_QUE_COMUNICAR "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_PARA_QUEM "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_UNID_ORGAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_QUANDO_OCORRE "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_RESP_CONTEUDO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_RESP_DIVULGACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_COMO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_APROVADOR_PB "						
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_DT_APROVACAO "
			
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO1 "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO2 "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO3 "
							
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " ) Values(" & intPlano2 & "," & intCdPCM & ",'" & strAtividade & "','" & strComunicacao & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'" & strOqueComunicar & "','" & strAQuemComunicar & "','" & strUnidadeOrgao & "','" & strQuandoOcorre & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'" & strRespConteudo & "','" & strRespDivulg & "','" & strComo & "','" & strAprovadorPB & "'"
			if strDtAprovacao <> "" then
				strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ",'" & strDtAprovacao & "',"
			else	
				strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", null ,"
			end if	
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'" & strfilArquivo1 & "','" & strfilArquivo2 & "','" & strfilArquivo3 & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
			'response.Write(strSQL_NovoPlanoPCM)
			'Response.End()
			on error resume next
				'if not blnNaoCadastraPlano then					
					'db_Cogest.Execute(strSQL_NovoPlanoGeral)
				'end if
				db_Cogest.Execute(strSQL_NovoPlanoPCM)
		
		'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH - INCLUSÃO
		elseif strPlano = "PAI" then		
		
			strCdInterface		= Trim(Ucase(Request("txtCdInterface")))		
			strGrupo 			= Trim(Ucase(Request("txtGrupo")))
			strTipoBatch 		= Request("selTipoBatch")
			strNomeInterface 	= Trim(Ucase(Request("txtNomeInterface")))		
			strPgrmEnvolv 		= Trim(Ucase(Request("txtPgrmEnvolv")))
			strPreRequisitos	= Trim(Ucase(Request("txtPreRequisitos")))			
			strRestricoes		= Trim(Ucase(Request("txtRestricoes")))			
			strDependencias		= Trim(Ucase(Request("txtDependencias")))	
			strRespAciona		= Trim(Ucase(Request("txtRespAciona")))	
			strProcedimento		= Trim(Ucase(Request("txtProcedimento")))	
			str_RespTecLegGeral	= Trim(Ucase(Request("txtRespTecLegGeral")))
			str_RespTecSinGeral	= Trim(Ucase(Request("txtRespTecSinGeral")))
								
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtInicio_Pai = 	split(Request("txtDtInicio_Pai"),"/")	
			strDia = vetDtInicio_Pai(0)
			strMes = vetDtInicio_Pai(1)
			strAno = vetDtInicio_Pai(2)	
			strDtInicio_Pai = strMes & "/" & strDia & "/" & strAno 
			
			Call VerificaUsuarioExistente("Responsável pelo Acionamento",strRespAciona, "Legado", "PAI")
			Call VerificaUsuarioExistente("Responsável Legado",str_RespTecLegGeral, "Legado", "PAI")			
			Call VerificaUsuarioExistente("Responsável Sinergia",str_RespTecSinGeral, "Sinergia", "PAI") 
								
			'Response.write strCdInterface & "<br>"
			'Response.write strGrupo & "<br>"
			'Response.write strTipoBatch & "<br>"
			'Response.write strNomeInterface 	& "<br>"
			'Response.write strPgrmEnvolv 	& "<br>"
			'Response.write strPreRequisitos	& "<br>"	
			'Response.write strRestricoes	& "<br>"
			'Response.write strDependencias 	& "<br>"
			'Response.write strRespAciona 	& "<br>"
			'Response.write strProcedimento 	& "<br>"
			'Response.write strDtInicio_Pai 	& "<br>"
			'Response.write str_RespTecLegGeral 	& "<br>"
			'Response.write str_RespTecSinGeral 	& "<br>"	
			'Response.end
																
			strSQL_NovoPlanoPAI = ""
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & " INSERT INTO XPEP_PLANO_TAREFA_PAI ( "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_CD_INTERFACE "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_GRUPO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_TIPO_PROCESSAMENTO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_NOME_INTERFACE "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_PROGRAMA_ENVOLVIDO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_PRE_REQUISITO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_RESTRICAO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_DEPENDENCIA "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_RESP_ACIONAMENTO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_DT_INICIO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_TX_PROCEDIMENTO "				
			'strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_NR_ID_PLANO_CONTINGENCIA "
			'strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", PPAI_NR_ID_PLANO_COMUNICACAO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", USUA_CD_USUARIO_RESP_SINER "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", USUA_CD_USUARIO_RESP_LEG "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & " ) Values(" & intPlano & "," & intCDPlanoGeral & ",'" & strCdInterface & "','" & strGrupo & "',"
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'" & strTipoBatch & "','" & strNomeInterface & "','" & strPgrmEnvolv & "','" & strPreRequisitos & "',"
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'" & strRestricoes & "','" & strDependencias & "','" & strRespAciona & "','" & strDtInicio_Pai & "','" & strProcedimento & "',"
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'" & str_RespTecSinGeral & "','" & str_RespTecLegGeral & "',"	
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
									
			'Response.write strSQL_NovoPlanoPAI
			'Response.end
			
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPAI)
				
		'*** PLANO DE AÇÕES CORRETIVAS E CONTINGÊNCIAS - INCLUSÃO
		elseif strPlano = "PAC" then				
		
			strProblemas			= Trim(Ucase(Request("txtProblemas")))	
			strAcoesCorrConting 	= Trim(Ucase(Request("txtAcoesCorrConting")))	
			strNomeInterface 		= Trim(Ucase(Request("txtNomeInterface")))					
			strUsuarioResponsavel 	= Trim(Ucase(Request("txtUsuarioResponsavel")))
			strRespTecSinGeral		= Trim(Ucase(Request("txtRespTecSinGeral")))			
			strRespFunSinGeral		= Trim(Ucase(Request("txtRespFunSinGeral")))										
			
			if Request("pIdAtividade") <> "" then
				intIdAtividade	= cint(Request("pIdAtividade"))
			else
				intIdAtividade	= 0
			end if
												
			'Response.write 	"FFF - " & intIdAtividade
			'Response.end							
												
			if trim(Request("txtDTAprovacao_PAC")) <> "" then					
				strDia = ""		
				strMes = ""
				strAno = ""
				vetDTAprovacao_PAC = split(Request("txtDTAprovacao_PAC"),"/")	
				strDia = vetDTAprovacao_PAC(0)
				strMes = vetDTAprovacao_PAC(1)
				strAno = vetDTAprovacao_PAC(2)	
				strDTAprovacao_PAC = strMes & "/" & strDia & "/" & strAno 
			else
				strDTAprovacao_PAC = ""
			end if
			
			Call VerificaUsuarioExistente("Responsável pelo Tratamento do Procedimento", strUsuarioResponsavel, "Legado", "PAC")					
			Call VerificaUsuarioExistente("Responsável Sinergia - Técnico", strRespTecSinGeral, "Sinergia", "PAC") 
			Call VerificaUsuarioExistente("Responsável Sinergia - Funcional", strRespFunSinGeral, "Sinergia", "PAC") 
			
			'Response.write strAcoesCorrConting & "<br>"		
			'Response.write strDTAprovacao_PAC	& "<br>"
			'Response.write strUsuarioResponsavel 	& "<br>"
			'Response.write strRespTecSinGeral 	& "<br>"
			'Response.write strRespFunSinGeral	& "<br>"			
			'Response.end
								
			'strSQL_NovoPlanoPAC = ""
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " INSERT INTO XPEP_PLANO_TAREFA_PAC ( "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " PLAN_NR_SEQUENCIA_PLANO "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PLTA_NR_SEQUENCIA_TAREFA "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_NR_ID_ATIVIDADE_PPO "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_TX_PROBLEMAS "			
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_TX_ACOES_CORR_CONT "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_DT_APROVACAO "		
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_TRAT_PROC "		
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_TEC "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_FUN "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_TX_OPERACAO "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_CD_NR_USUARIO "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_DT_ATUALIZACAO "
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " ) Values(" & intPlano & "," & intCDPlanoGeral & "," & intIdAtividade & ",'" & strProblemas & "','" & strAcoesCorrConting & "',"
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & "'" & strDTAprovacao_PAC & "','" & strUsuarioResponsavel & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
			'strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
										
										
			'*** VERIFICA SE O PLANO JÁ FOI CADASTRADO EM XPEP_PLANO_TAREFA_PAC
			str_sqlAtividade = ""
			str_sqlAtividade = str_sqlAtividade & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
			str_sqlAtividade = str_sqlAtividade & ", PLTA_NR_SEQUENCIA_TAREFA "					
			str_sqlAtividade = str_sqlAtividade & " FROM XPEP_PLANO_TAREFA_PAC"
			str_sqlAtividade = str_sqlAtividade & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
			str_sqlAtividade = str_sqlAtividade & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
			'Response.write str_sqlAtividade
			'Response.end
			set rdsVerificaAtividadePAC = db_Cogest.Execute(str_sqlAtividade)	
			
			if rdsVerificaAtividadePAC.eof then
				'*** GRAVA O PAC NA TABELA XPEP_PLANO_TAREFA_PAC								
				strSQL_NovoPlanoPAC = ""
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " INSERT INTO XPEP_PLANO_TAREFA_PAC ( "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " PLAN_NR_SEQUENCIA_PLANO "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PLTA_NR_SEQUENCIA_TAREFA "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_NR_ID_ATIVIDADE_PPO	"		
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_TX_OPERACAO "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_CD_NR_USUARIO "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_DT_ATUALIZACAO "
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " ) Values(" & intPlano & "," & intCDPlanoGeral & ",0," 
				strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & "'I','" & Session("CdUsuario") & "',GETDATE())"
				'Response.write strSQL_NovoPlanoPAC
				'Response.end	
				db_Cogest.Execute(strSQL_NovoPlanoPAC)				
			end if						
			rdsVerificaAtividadePAC.close
			set rdsVerificaAtividadePAC = nothing									
																								
			'*** Seleciona o cod para a Nova Seq para Sub-Atividade de PDS
			intCdSeqFunc = 0	
			str_SQL = ""
			str_SQL = str_SQL & " SELECT "
			str_SQL = str_SQL & " MAX(PPAC_NR_SEQUENCIA_FUNC) AS int_Max_SeqFunc "
			str_SQL = str_SQL & " FROM XPEP_PLANO_TAREFA_PAC_FUNC "
			str_SQL = str_SQL & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
			str_SQL = str_SQL & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
			Set rdsMaxFunc = db_Cogest.Execute(str_SQL)			
			if isnull(rdsMaxFunc("int_Max_SeqFunc")) then
				intCdSeqFunc = 1
			else
				intCdSeqFunc = rdsMaxFunc("int_Max_SeqFunc") + 1
			end if
			rdsMaxFunc.Close	
			set rdsMaxFunc = nothing																					
														
			'Response.write 	intCdSeqFunc
			'Response.end												
																								
			strSQL_NovoPlanoPAC_Sub = ""
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & " INSERT INTO XPEP_PLANO_TAREFA_PAC_FUNC ( "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PPAC_NR_SEQUENCIA_FUNC "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PPAC_NR_ID_ATIVIDADE_PPO "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PPAC_TX_PROBLEMAS "			
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PPAC_TX_ACOES_CORR_CONT "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", PPAC_DT_APROVACAO "		
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", USUA_CD_USUARIO_RESP_TRAT_PROC "		
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", USUA_CD_USUARIO_RESP_SIN_TEC "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", USUA_CD_USUARIO_RESP_SIN_FUN "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & " ) Values(" & intPlano & "," & intCDPlanoGeral & "," & intCdSeqFunc & "," & intIdAtividade & ",'" & strProblemas & "','" & strAcoesCorrConting & "',"
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & "'" & strDTAprovacao_PAC & "','" & strUsuarioResponsavel & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
			strSQL_NovoPlanoPAC_Sub = strSQL_NovoPlanoPAC_Sub & "'I','" & Session("CdUsuario") & "',GETDATE())" 																	
								
			'Response.write 	strSQL_NovoPlanoPAC_Sub
			'Response.end 
										
			on error resume next					
				db_Cogest.Execute(strSQL_NovoPlanoPAC_Sub)
		
		'*** PLANO DE CONVERSÕES DE DADOS - INCLUSÃO
		elseif strPlano = "PCD" then		
					
			strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
			strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))
			strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
			strRespFunSinGeral 	= Trim(Ucase(Request("txtRespFunSinGeral")))
			
			strDesenvAssociados	= Request("pSistemas")
			strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))					
			strSistLegado 		= Request("txtSistLegado")			
			strTipoCarga		= Request("selTipoCarga")
			strTipoDados 		= Request("selTipoDados")
			strCaractDado 		= Request("selCaractDado")
			intExtracao_PCD		= Request("txtExtracao_PCD")
			strExtracao_Unid	= Request("txtExtracao_Unid")
			intCarga_PCD	 	= Request("txtCarga_PCD")
			strCarga_Unid		= Request("txtCarga_Unid")
			strArqCarga		 	= Trim(Ucase(Request("txtArqCarga")))			
			intVolume 			= Request("txtVolume")
			strDependencias		= Trim(Ucase(Request("txtDependencias")))			
			strComoExecuta		= Trim(Ucase(Request("txtComoExecuta")))		
			
			if Request("txtDTExtracao_PCD") <> "" then			
				strDia = ""		
				strMes = ""
				strAno = ""
				vetDTExtracao = split(Request("txtDTExtracao_PCD"),"/")	
				strDia = vetDTExtracao(0)
				strMes = vetDTExtracao(1)
				strAno = vetDTExtracao(2)	
				strDTExtracao_PCD = strMes & "/" & strDia & "/" & strAno 
			else
				strDTExtracao_PCD = ""
			end if
						
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTCargaIni = 	split(Request("txtDTCarga_PCD_Inicio"),"/")	
			strDia = vetDTCargaIni(0)
			strMes = vetDTCargaIni(1)
			strAno = vetDTCargaIni(2)	
			srtDTCarga_PCD_Ini = strMes & "/" & strDia & "/" & strAno 
						
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTCargaFim = 	split(Request("txtDTCarga_PCD_Fim"),"/")	
			strDia = vetDTCargaFim(0)
			strMes = vetDTCargaFim(1)
			strAno = vetDTCargaFim(2)	
			srtDTCarga_PCD_Fim = strMes & "/" & strDia & "/" & strAno 																			
															
			if strRespTecLegGeral <> "" then												
				Call VerificaUsuarioExistente("Responsável Legado - Técnico",strRespTecLegGeral, "Legado", "PCD")	
			end if
			
			if strRespFunLegGeral <> "" then	
				Call VerificaUsuarioExistente("Responsável Legado - Funcional",strRespFunLegGeral, "Legado", "PCD")			
			end if
			
			if strRespTecSinGeral <> "" then
				Call VerificaUsuarioExistente("Responsável Sinergia - Técnico",strRespTecSinGeral, "Sinergia", "PCD")
			end if
			
			if strRespFunSinGeral <> "" then 
				Call VerificaUsuarioExistente("Responsável Sinergia - Funcional",strRespFunSinGeral, "Sinergia", "PCD")
			end if
			
			'Response.write strRespTecLegGeral & "<br>"
			'Response.write strRespFunLegGeral & "<br>"
			'Response.write strRespTecSinGeral & "<br>"
			'Response.write strRespFunSinGeral & "<br>"			
			'Response.write strDesenvAssociados & "<br>"
			'Response.write strDadoMigrado	& "<br>"
			'Response.write strSistLegado & "<br>"
			'Response.write strTipoCarga	& "<br>"
			'Response.write strTipoDados & "<br>"
			'Response.write strCaractDado & "<br>"
			'Response.write intExtracao_PCD	& "<br>"
			'Response.write strExtracao_Unid	& "<br>"			
			'Response.write intCarga_PCD	 & "<br>"
			'Response.write strCarga_Unid & "<br>"			
			'Response.write strArqCarga	& "<br>"		
			'Response.write intVolume & "<br>"
			'Response.write strDependencias	& "<br>"	
			'Response.write strComoExecuta	& "<br>"		
			'Response.write strDTExtracao_PCD & "<br>"
			'Response.write srtDTCarga_PCD_Ini & "<br>"
			'Response.write srtDTCarga_PCD_Fim		
			'Response.end																						
												
			strSQL_NovoPlanoPCD = ""
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & " INSERT INTO XPEP_PLANO_TAREFA_PCD ("
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PLTA_NR_SEQUENCIA_TAREFA "		
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_SISTEMA_LEGADO "					
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_DADO_A_SER_MIGRADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_TIPO_ATIVIDADE "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_TIPO_DADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_CARAC_DADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_ARQ_CARGA "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_NR_VOLUME "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_DEPENDENCIAS "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_DT_EXTRACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_DT_CARGA_INICIO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_DT_CARGA_FIM "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_tx_COMO_EXECUTA "						
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_TEC "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ") Values(" & intPlano & ","
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & intCDPlanoGeral & ",'" & strSistLegado & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strDadoMigrado & "','" & strTipoCarga & "','" & strTipoDados & "','" & strCaractDado & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & intExtracao_PCD & ",'" & strExtracao_Unid & "'," & intCarga_PCD & ",'" & strCarga_Unid & "','" & strArqCarga & "'," & intVolume & ","
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strDependencias & "','" & strDTExtracao_PCD & "','" & srtDTCarga_PCD_Ini & "','" & srtDTCarga_PCD_Fim & "','" & strComoExecuta & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strRespTecLegGeral & "','" & strRespFunLegGeral & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
				
			'Response.write strSQL_NovoPlanoPCD		
			'Response.end	
				
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPCD)	
				
				'*** GRAVAÇÃO DOS DESENVOLVIMENTOS ASSOCIADOS PARA ESTE PLANO ***
				if strDesenvAssociados <> "" then				
					i = 0
					vetDesenvAssociados = split(strDesenvAssociados,"|")				
					for i = lbound(vetDesenvAssociados) to ubound(vetDesenvAssociados) 
						if vetDesenvAssociados(i) <> "" then
							sql_NovoDesenvAss = ""
							sql_NovoDesenvAss = sql_NovoDesenvAss & " INSERT INTO XPEP_TAREFA_DESENVOLVIMENTO ("
							sql_NovoDesenvAss = sql_NovoDesenvAss & " PLAN_NR_SEQUENCIA_PLANO"
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", PLTA_NR_SEQUENCIA_TAREFA"
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", PPCD_NR_SEQUENCIA_FUNC"
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", DESE_CD_DESENVOLVIMENTO "
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_TX_OPERACAO "
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_CD_NR_USUARIO "
							sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_DT_ATUALIZACAO"
							sql_NovoDesenvAss = sql_NovoDesenvAss & ") Values(" & intPlano & ","
							sql_NovoDesenvAss = sql_NovoDesenvAss & intCDPlanoGeral & ",'0','" & vetDesenvAssociados(i) & "',"
							sql_NovoDesenvAss = sql_NovoDesenvAss & "'I','" & Session("CdUsuario") & "',GETDATE())" 	
							'Response.write sql_NovoDesenvAss
							'Response.end	
							'db_Cogest.Execute(sql_NovoDesenvAss)						
						end if					
					next
				end if		
				
		'*** PLANO DE DESLIGAMENTO DE SISTEMAS LEGADOS - INCLUSÃO
		elseif strPlano = "PDS" then		

			intSistLegado 		= Request("selSistLegado")					
			strRespTecLeg 		= Trim(Ucase(Request("txtRespTecLeg")))
			strRespFunLeg 		= Trim(Ucase(Request("txtRespFunLeg")))
			strTpDesligamento 	= Request("rdbTpDesligamento")
			strGerTecRespLeg 	= Request("txtGerTecRespLeg")
						
			'Response.write intPlano & "<br>"
			'Response.write intCDPlanoGeral & "<br>"
			'Response.write intSistLegado & "<br>"
			'Response.write strRespTecLeg & "<br>"
			'Response.write strRespFunLeg & "<br>"
			'Response.write strTpDesligamento & "<br>"
			'Response.write strGerTecRespLeg	& "<br>"			
			'Response.end	
			
			Call VerificaUsuarioExistente("Responsável Legado - Técnico", strRespTecLeg, "Legado", "PDS")			
			Call VerificaUsuarioExistente("Responsável Legado - Funcional", strRespFunLeg, "Legado", "PDS") 									

			int_SeqCDTarefa = intCDPlanoGeral
												
			strSQL_NovoPlanoPDS = ""
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " INSERT INTO XPEP_PLANO_TAREFA_PDS ( "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", SIST_NR_SEQUENCIAL_SISTEMA_LEGADO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", USUA_CD_USUARIO_RESP_LEG_TEC "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", USUA_CD_USUARIO_RESP_LEG_FUN "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", PPDS_TX_TIPO_DESLIGAMENTO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", PPDS_TX_GER_TEC_RESP_LEGADO "									
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ) Values(" 
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " " & intPlano 
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ," & int_SeqCDTarefa
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ," & intSistLegado 
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ,'" & strRespTecLeg & "'"
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ,'" & strRespFunLeg & "'"
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ,'" & strTpDesligamento & "'"
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ,'" & strGerTecRespLeg & "'"			
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & ",'I','" & Session("CdUsuario") & "',GETDATE())" 		
				
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPDS)		
								
		end if
	
		if err.number = 0 then
		
			strMSG = "Detalhamento incluido com sucesso."
		    strGravado = 1
			'set correio = server.CreateObject("Persits.MailSender")
			'correio.host = "harpia.petrobras.com.br"
			 
			'correio.from="cursos@xproc.com"
		
			'correio.AddAddress "robson_28.infotec@petrobras.com.br"	     			
			'correio.AddAddress "gustavogomes.bearingpoint@petrobras.com.br"
			'correio.AddAddress "sergio.salomao.bearingpoint@petrobras.com.br"
							
			'correio.Subject="Inclusão de Novo Curso"
						
			'data_Atual=day(date) &"/"& month(date) &"/"& year(date)
							
			'correio.Body=" O Curso '" & UCASE(valor_cod) & "' foi INCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
			'correio.send
		else
			strMSG = "Houve um erro no cadastro do detalhamento."
		end if	 
	end if	
	rdsVerificaAtvidade.close
	set rdsVerificaAtvidade = nothing

'************************************** ALTERAÇÃO ************************************************	
elseif strAcao = "A" then

	if strPlano <> "PCM" then
		strVerificaAtvidade = ""
		strVerificaAtvidade = strVerificaAtvidade & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
		strVerificaAtvidade = strVerificaAtvidade & " FROM XPEP_PLANO_TAREFA_GERAL"
		strVerificaAtvidade = strVerificaAtvidade & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & intIdTaskProject
		Set rdsVerificaAtvidade = db_Cogest.Execute(strVerificaAtvidade)		
		
		if not rdsVerificaAtvidade.EOF then
			int_SeqCDTarefa = cint(rdsVerificaAtvidade("PLTA_NR_SEQUENCIA_TAREFA"))
		else	
			int_SeqCDTarefa = 999999
		end if
	end if
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	if strPlano = "PPO" then
				
		txtDescrParada 		= Trim(Ucase(Request("txtDescrParada")))			
		strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
		strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
		strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))			
		intTempParada 		= Request("txtTempParada")
		strUnidadeMedida	= Request("selUnidMedida")			
		strProcedParada 	= Trim(Ucase(Request("txtProcedParada")))			
		strUsuarioGestor 	= Trim(Ucase(Request("txtUsuarioGestor")))					
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataParada = 	split(Request("txtDtParadaLegado"),"/")	
		strDia = vetDtDataParada(0)
		strMes = vetDtDataParada(1)
		strAno = vetDtDataParada(2)	
		strDtParadaLegado = strMes & "/" & strDia & "/" & strAno 
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataR3 = 	split(Request("txtDtIniR3"),"/")	
		strDia = vetDtDataR3(0)
		strMes = vetDtDataR3(1)
		strAno = vetDtDataR3(2)	
		strDtIniR3 = strMes & "/" & strDia & "/" & strAno 
							
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtLimiteAprov = 	split(Request("txtDtLimiteAprov"),"/")	
		strDia = vetDtLimiteAprov(0)
		strMes = vetDtLimiteAprov(1)
		strAno = vetDtLimiteAprov(2)	
		strDtLimiteAprov = strMes & "/" & strDia & "/" & strAno 
		
		Call VerificaUsuarioExistente("Responsável Sinergia",strRespTecSinGeral, "Sinergia", "PPO")			
		Call VerificaUsuarioExistente("Responsável Legado - Técnico",strRespTecLegGeral, "Legado", "PPO")
		Call VerificaUsuarioExistente("Responsável Legado - Funcional",strRespFunLegGeral, "Legado", "PPO")
		Call VerificaUsuarioExistente("Gestor do Processo",strUsuarioGestor, "Legado", "PPO")
		
		'Response.write int_SeqCDTarefa & "<br>"			
		'Response.write intIdTaskProject & "<br>"
		'Response.write intPlano & "<br>"
		'Response.write txtDescrParada & "<br>"
		'Response.write strRespTecSinGeral & "<br>"
		'response.write strRespTecLegGeral & "<br>"
		'Response.write strRespFunLegGeral & "<br>"
		'response.write intTempParada & "<br>"
		'Response.write strUnidadeMedida & "<br>"	
		'Response.write strProcedParada & "<br>"
		'response.write strDtParadaLegado& "<br>"
		'Response.write strDtIniR3 & "<br>"
		'response.write strUsuarioGestor & "<br>"
		'Response.write strDtLimiteAprov & "<br><br>"
		'Response.end	
											
		strSQL_AltPlanoPPO = ""
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & " UPDATE XPEP_PLANO_TAREFA_PPO SET"		
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & " PPPO_TX_DESCRICAO_PARADA = '" & txtDescrParada & "'"	
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_TX_QTD_TEMPO_PARADA = " & intTempParada
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_TX_UNID_TEMPO_PARADA = '" & strUnidadeMedida & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_TX_PROCEDIMENTOS_PARADA ='" & strProcedParada & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_DT_PARADA_LEGADO = '" & strDtParadaLegado & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_DT_INICIO_R3 = '" & strDtIniR3 & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_DT_LIMITE_APROVACAO = '" & strDtLimiteAprov & "'"
		'strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_NR_ID_PLANO_CONTINGENCIA "
		'strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", PPPO_NR_ID_PLANO_COMUNICACAO "	
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", USUA_CD_USUARIO_RESP_SINER = '" & strRespTecSinGeral & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", USUA_CD_USUARIO_RESP_LEG_TEC = '" & strRespTecLegGeral & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", USUA_CD_USUARIO_RESP_LEG_FUN = '" & strRespFunLegGeral & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", USUA_CD_USUARIO_GESTOR_PROC = '" & strUsuarioGestor & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 
			
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPPO)			
		
	'*** PLANO DE COMUNICAÇÃO- ALTERAÇÃO
	elseif strPlano = "PCM" then

		str_CdSeqPCM 		= request("pCdSeqPCM")
		strAtividade 		= Trim(Ucase(Request("txtAtividade")))									
		strComunicacao		= Request("selComunicacao")			
		strOqueComunicar 	= Trim(Ucase(Request("txtOqueComunicar")))
		strAQuemComunicar 	= Trim(Ucase(Request("txtAQuemComunicar")))
		strUnidadeOrgao 	= Trim(Ucase(Request("txtUnidadeOrgao")))			
		strRespConteudo		= Trim(Ucase(Request("txtRespConteudo")))		
		strRespDivulg		= Trim(Ucase(Request("txtRespDivulg")))
		strComo				= Trim(Ucase(Request("txtComo")))
		strAprovadorPB		= Trim(Ucase(Request("txtAprovadorPB")))	
		strfilArquivo1		= Trim(Ucase(Request("filArquivo1")))		
		strfilArquivo2		= Trim(Ucase(Request("filArquivo2")))		
		strfilArquivo3		= Trim(Ucase(Request("filArquivo3")))		
		
		strDia = ""		
		strMes = ""
		strAno = ""			
		vetDtQuandoOcorre = split(Request("txtQuandoOcorre"),"/")	
		strDia = vetDtQuandoOcorre(0)
		strMes = vetDtQuandoOcorre(1)
		strAno = vetDtQuandoOcorre(2)	
		strQuandoOcorre = strMes & "/" & strDia & "/" & strAno 
													
		if trim(Request("txtDtAprovacao")) <> "" then					
			strDia = ""		
			strMes = ""
			strAno = ""			
			vetDtAprovacao = 	split(Request("txtDtAprovacao"),"/")	
			strDia = vetDtAprovacao(0)
			strMes = vetDtAprovacao(1)
			strAno = vetDtAprovacao(2)	
			strDtAprovacao = strMes & "/" & strDia & "/" & strAno 
		end if
		
		if strAprovadorPB <> "" then
			Call VerificaUsuarioExistente("Aprovador PB",strAprovadorPB, "Legado", "PCM")	
		end if
		
		'Response.write strComunicacao & "<br>"
		'Response.write strOqueComunicar & "<br>"
		'Response.write strAQuemComunicar & "<br>"
		'Response.write strUnidadeOrgao 	& "<br>"
		'Response.write strQuandoOcorre 	& "<br>"
		'Response.write strRespConteudo	& "<br>"
		'Response.write strRespDivulg	& "<br>"
		'Response.write strAprovadorPB	& "<br>"
		'Response.write strDtAprovacao 	& "<br><br><br>"
		'Response.end
													
		strSQL_AltPlanoPCM = ""
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " UPDATE XPEP_PLANO_TAREFA_PCM SET"	
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " PPCM_TX_ATIVIDADE = '" & strAtividade & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_TP_COMUNICACAO = '" & strComunicacao & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_O_QUE_COMUNICAR = '" & strOqueComunicar & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_PARA_QUEM = '" & strAQuemComunicar & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_UNID_ORGAO = '" & strUnidadeOrgao & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_QUANDO_OCORRE = '" & strQuandoOcorre & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_RESP_CONTEUDO = '" & strRespConteudo & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_RESP_DIVULGACAO = '" & strRespDivulg & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_COMO = '" & strComo & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_APROVADOR_PB = '" & strAprovadorPB & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_DT_APROVACAO = '" & strDtAprovacao & "'"	
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"			
		if strfilArquivo1 <> "" then
			strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO1 = '" & strfilArquivo1 & "'"	
		end if 		
		if strfilArquivo2 <> "" then
			strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO2 = '" & strfilArquivo2 & "'"	
		end if		
		if strfilArquivo3 <> "" then
			strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_ARQUIVO_ANEXO3 = '" & strfilArquivo3 & "'"	
		end if			
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano2
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " AND PPCM_NR_SEQUENCIA_TAREFA = " & str_CdSeqPCM

		'Response.write strSQL_AltPlanoPCM
		'Response.end		
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCM)	
	
	'*** PLANO DE AÇÕES CORRETIVAS E CONTINGÊNCIAS - ALTERAÇÃO
	elseif strPlano = "PAC" then		
	
		'****** PEGARÁ O COD DA TAREFA LÁ DO PROJECT
		intCDPlanoGeral = intIdTaskProject
	
		strProblemas			= Trim(Ucase(Request("txtProblemas")))	
		strAcoesCorrConting 	= Trim(Ucase(Request("txtAcoesCorrConting")))	
		strNomeInterface 		= Trim(Ucase(Request("txtNomeInterface")))	
		strUsuarioResponsavel 	= Trim(Ucase(Request("txtUsuarioResponsavel")))
		strRespTecSinGeral		= Trim(Ucase(Request("txtRespTecSinGeral")))			
		strRespFunSinGeral		= Trim(Ucase(Request("txtRespFunSinGeral")))										
											
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDTAprovacao_PAC = 	split(Request("txtDTAprovacao_PAC"),"/")	
		strDia = vetDTAprovacao_PAC(0)
		strMes = vetDTAprovacao_PAC(1)
		strAno = vetDTAprovacao_PAC(2)	
		strDTAprovacao_PAC = strMes & "/" & strDia & "/" & strAno 
		
		Call VerificaUsuarioExistente("Responsável pelo Tratamento do Procedimento", strUsuarioResponsavel, "Legado", "PAC")					
		Call VerificaUsuarioExistente("Responsável Sinergia - Técnico", strRespTecSinGeral, "Sinergia", "PAC") 
		Call VerificaUsuarioExistente("Responsável Sinergia - Funcional", strRespFunSinGeral, "Sinergia", "PAC") 
					
		intCdSeqPAC = Request("pCdSeqPAC")
		
		'Response.write intCdSeqPAC & "<br>"							
		'Response.write strAcoesCorrConting & "<br>"		
		'Response.write strDTAprovacao_PAC	& "<br>"
		'Response.write strUsuarioResponsavel 	& "<br>"
		'Response.write strRespTecSinGeral 	& "<br>"
		'Response.write strRespFunSinGeral	& "<br>"			
		'Response.end
								
		strSQL_AltPlanoPAC = ""
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " UPDATE XPEP_PLANO_TAREFA_PAC_FUNC SET"			
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " PPAC_TX_PROBLEMAS = '" & strProblemas & "'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", PPAC_TX_ACOES_CORR_CONT = '" & strAcoesCorrConting & "'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", PPAC_DT_APROVACAO = '" & strDTAprovacao_PAC & "'"		
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_TRAT_PROC = '" & strUsuarioResponsavel & "'"		
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_TEC = '" & strRespTecSinGeral & "'"	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_FUN = '" & strRespFunSinGeral & "'"	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " AND PPAC_NR_SEQUENCIA_FUNC = " & intCdSeqPAC							
									
		'Response.write strSQL_AltPlanoPAC
		'Response.end
		
		on error resume next			
			db_Cogest.Execute(strSQL_AltPlanoPAC)
	
	
	'*** PLANO DE CONVERSÕES DE DADOS - ALTERAÇÃO
	elseif strPlano = "PCD" then
								
		strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
		strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))
		strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
		strRespFunSinGeral 	= Trim(Ucase(Request("txtRespFunSinGeral")))
		
		strDesenvAssociados	= Request("pSistemas")
		strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))					
		strSistLegado 		= Request("txtSistLegado")			
		strTipoCarga		= Request("selTipoCarga")
		strTipoDados 		= Request("selTipoDados")
		strCaractDado 		= Request("selCaractDado")
		intExtracao_PCD		= Request("txtExtracao_PCD")
		strExtracao_Unid	= Request("txtExtracao_Unid")
		intCarga_PCD	 	= Request("txtCarga_PCD")
		strCarga_Unid		= Request("txtCarga_Unid")
		strArqCarga		 	= Trim(Ucase(Request("txtArqCarga")))			
		intVolume 			= Request("txtVolume")
		strDependencias		= Trim(Ucase(Request("txtDependencias")))			
		strComoExecuta		= Trim(Ucase(Request("txtComoExecuta")))		
					
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDTExtracao = 	split(Request("txtDTExtracao_PCD"),"/")	
		strDia = vetDTExtracao(0)
		strMes = vetDTExtracao(1)
		strAno = vetDTExtracao(2)	
		strDTExtracao_PCD = strMes & "/" & strDia & "/" & strAno 
		
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDTCargaIni = 	split(Request("txtDTCarga_PCD_Inicio"),"/")	
		strDia = vetDTCargaIni(0)
		strMes = vetDTCargaIni(1)
		strAno = vetDTCargaIni(2)	
		srtDTCarga_PCD_Ini = strMes & "/" & strDia & "/" & strAno 
		
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDTCargaFim = 	split(Request("txtDTCarga_PCD_Fim"),"/")	
		strDia = vetDTCargaFim(0)
		strMes = vetDTCargaFim(1)
		strAno = vetDTCargaFim(2)	
		srtDTCarga_PCD_Fim = strMes & "/" & strDia & "/" & strAno 
		
		if strRespTecLegGeral <> "" then												
			Call VerificaUsuarioExistente("Responsável Legado - Técnico",strRespTecLegGeral, "Legado", "PCD")	
		end if
		
		if strRespFunLegGeral <> "" then	
			Call VerificaUsuarioExistente("Responsável Legado - Funcional",strRespFunLegGeral, "Legado", "PCD")			
		end if
		
		if strRespTecSinGeral <> "" then
			Call VerificaUsuarioExistente("Responsável Sinergia - Técnico",strRespTecSinGeral, "Sinergia", "PCD")
		end if
		
		if strRespFunSinGeral <> "" then 
			Call VerificaUsuarioExistente("Responsável Sinergia - Funcional",strRespFunSinGeral, "Sinergia", "PCD")
		end if
			
		'Call VerificaUsuarioExistente("Responsável Legado - Técnico",strRespTecLegGeral, "Legado", "PCD")	
		'Call VerificaUsuarioExistente("Responsável Legado - Funcional",strRespFunLegGeral, "Legado", "PCD")			
		'Call VerificaUsuarioExistente("Responsável Sinergia - Técnico",strRespTecSinGeral, "Sinergia", "PCD") 
		'Call VerificaUsuarioExistente("Responsável Sinergia - Funcional",strRespFunSinGeral, "Sinergia", "PCD")
		
		'Response.write strRespTecLegGeral & "<br>"
		'Response.write strRespFunLegGeral & "<br>"				
		'Response.write strRespTecSinGeral & "<br>"
		'Response.write strRespFunSinGeral & "<br>"
		'Response.write strDesenvAssociados & "<br>"
		'Response.write strDadoMigrado	& "<br>"
		'Response.write strSistLegado & "<br>"
		'Response.write strTipoCarga	& "<br>"
		'Response.write strTipoDados & "<br>"
		'Response.write strCaractDado & "<br>"
		'Response.write intExtracao_PCD	& "<br>"
		'Response.write strExtracao_Unid	& "<br>"		
		'Response.write intCarga_PCD	 & "<br>"
		'Response.write strCarga_Unid	& "<br>"
		'Response.write strArqCarga	& "<br>"		
		'Response.write intVolume & "<br>"
		'Response.write strDependencias	& "<br>"	
		'Response.write strComoExecuta	& "<br>"		
		'Response.write strDTExtracao_PCD & "<br>"
		'Response.write srtDTCarga_PCD_Ini & "<br>"
		'Response.write srtDTCarga_PCD_Fim	& "<br><br><br>"
		'Response.end													
													
		strSQL_AltPlanoPCD = ""
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " UPDATE XPEP_PLANO_TAREFA_PCD SET"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " PPCD_TX_SISTEMA_LEGADO = '" & strSistLegado & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_DADO_A_SER_MIGRADO = '" & strDadoMigrado & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_TIPO_ATIVIDADE = '" & strTipoCarga & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_TIPO_DADO = '" & strTipoDados & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_CARAC_DADO = '" & strCaractDado & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO = " & intExtracao_PCD		
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO = '" & strExtracao_Unid & "'"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA = " & intCarga_PCD
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA = '" & strCarga_Unid & "'"			
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_ARQ_CARGA = '" & strArqCarga & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_VOLUME = " & intVolume
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_DEPENDENCIAS = '" & strDependencias & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_DT_EXTRACAO = '" & strDTExtracao_PCD & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_DT_CARGA_INICIO = '" & srtDTCarga_PCD_Ini & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_DT_CARGA_FIM = '" & srtDTCarga_PCD_Fim & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_COMO_EXECUTA = '" & strComoExecuta & "'"
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_TEC = '" & strRespTecLegGeral & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_FUN = '" & strRespFunLegGeral & "'"			
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_TEC = '" & strRespTecSinGeral & "'"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_FUN = '" & strRespFunSinGeral & "'"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 			
		'Response.write strSQL_AltPlanoPCD
		'Response.end	
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCD)	
			
			'*** EXCLUSÃO DOS DESENVOLVIMENTOS ASSOCIADOS PARA ESTE PLANO ***				
			sql_DelDesenvAss = ""
			sql_DelDesenvAss = sql_DelDesenvAss & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"
			sql_DelDesenvAss = sql_DelDesenvAss & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano 
			sql_DelDesenvAss = sql_DelDesenvAss & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa		
			sql_DelDesenvAss = sql_DelDesenvAss & " AND PPCD_NR_SEQUENCIA_FUNC = '0'"		
			
			'Response.write sql_DelDesenvAss
			'Response.end
			db_Cogest.Execute(sql_DelDesenvAss)						
						
			'*** GRAVAÇÃO DOS DESENVOLVIMENTOS ASSOCIADOS PARA ESTE PLANO ***
			if strDesenvAssociados <> "" then				
				i = 0
				vetDesenvAssociados = split(strDesenvAssociados,"|")				
				for i = lbound(vetDesenvAssociados) to ubound(vetDesenvAssociados) 
					if vetDesenvAssociados(i) <> "" then
						sql_NovoDesenvAss = ""
						sql_NovoDesenvAss = sql_NovoDesenvAss & " INSERT INTO XPEP_TAREFA_DESENVOLVIMENTO ("
						sql_NovoDesenvAss = sql_NovoDesenvAss & " PLAN_NR_SEQUENCIA_PLANO"
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", PLTA_NR_SEQUENCIA_TAREFA"
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", PPCD_NR_SEQUENCIA_FUNC"
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", DESE_CD_DESENVOLVIMENTO "
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_TX_OPERACAO "
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_CD_NR_USUARIO "
						sql_NovoDesenvAss = sql_NovoDesenvAss & ", ATUA_DT_ATUALIZACAO"
						sql_NovoDesenvAss = sql_NovoDesenvAss & ") Values(" & intPlano & ","
						sql_NovoDesenvAss = sql_NovoDesenvAss & int_SeqCDTarefa & ",'0','" & vetDesenvAssociados(i) & "',"
						sql_NovoDesenvAss = sql_NovoDesenvAss & "'I','" & Session("CdUsuario") & "',GETDATE())"
						db_Cogest.Execute(sql_NovoDesenvAss)						
					end if					
				next
			end if		
			
	'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH - ALTERAÇÃO
	elseif strPlano = "PAI" then								
		
		strCdInterface		= Trim(Ucase(Request("txtCdInterface")))		
		strGrupo 			= Trim(Ucase(Request("txtGrupo")))
		strTipoBatch 		= Request("selTipoBatch")
		strNomeInterface 	= Trim(Ucase(Request("txtNomeInterface")))		
		strPgrmEnvolv 		= Trim(Ucase(Request("txtPgrmEnvolv")))
		strPreRequisitos	= Trim(Ucase(Request("txtPreRequisitos")))			
		strRestricoes		= Trim(Ucase(Request("txtRestricoes")))			
		strDependencias		= Trim(Ucase(Request("txtDependencias")))	
		strRespAciona		= Trim(Ucase(Request("txtRespAciona")))	
		strProcedimento		= Trim(Ucase(Request("txtProcedimento")))	
		str_RespTecLegGeral	= Trim(Ucase(Request("txtRespTecLegGeral")))
		str_RespTecSinGeral	= Trim(Ucase(Request("txtRespTecSinGeral")))
								
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtInicio_Pai = 	split(Request("txtDtInicio_Pai"),"/")	
		strDia = vetDtInicio_Pai(0)
		strMes = vetDtInicio_Pai(1)
		strAno = vetDtInicio_Pai(2)	
		strDtInicio_Pai = strMes & "/" & strDia & "/" & strAno
		
		Call VerificaUsuarioExistente("Responsável pelo Acionamento",strRespAciona, "Legado", "PAI")
		Call VerificaUsuarioExistente("Responsável Legado",str_RespTecLegGeral, "Legado", "PAI")			
		Call VerificaUsuarioExistente("Responsável Sinergia",str_RespTecSinGeral, "Sinergia", "PAI") 
			
		'Response.write strCdInterface & "<br>"
		'Response.write strGrupo & "<br>"
		'Response.write strTipoBatch & "<br>"
		'Response.write strNomeInterface 	& "<br>"
		'Response.write strPgrmEnvolv 	& "<br>"
		'Response.write strPreRequisitos	& "<br>"	
		'Response.write strRestricoes	& "<br>"
		'Response.write strDependencias 	& "<br>"
		'Response.write strRespAciona 	& "<br>"
		'Response.write strProcedimento 	& "<br>"
		'Response.write strDtInicio_Pai 	& "<br>"
		'Response.write str_RespTecLegGeral 	& "<br>"
		'Response.write str_RespTecSinGeral 	& "<br><br><br>"	
		'Response.end
													
		strSQL_AltPlanoPAI = ""
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & " UPDATE XPEP_PLANO_TAREFA_PAI SET"	
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & " PPAI_TX_CD_INTERFACE = '" & strCdInterface & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_GRUPO = '" & strGrupo & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_TIPO_PROCESSAMENTO = '" & strTipoBatch & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_NOME_INTERFACE = '" & strNomeInterface & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_PROGRAMA_ENVOLVIDO = '" & strPgrmEnvolv & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_PRE_REQUISITO = '" & strPreRequisitos & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_RESTRICAO = '" & strRestricoes & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_DEPENDENCIA = '" & strDependencias & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_RESP_ACIONAMENTO = '" & strRespAciona & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_DT_INICIO = '" & strDtInicio_Pai & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_TX_PROCEDIMENTO = '" & strProcedimento & "'"		
		'strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_NR_ID_PLANO_CONTINGENCIA "
		'strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", PPAI_NR_ID_PLANO_COMUNICACAO "
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", USUA_CD_USUARIO_RESP_SINER = '" & str_RespTecSinGeral & "'"	
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", USUA_CD_USUARIO_RESP_LEG = '" & str_RespTecLegGeral & "'"	
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPAI = strSQL_AltPlanoPAI & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 
		
		'Response.write strSQL_AltPlanoPAI
		'Response.end		
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPAI)	

	'*** PLANO DE DESLIGAMENTO DE SISTEMAS LEGADOS - ALTERAÇÃO
	elseif strPlano = "PDS" then								
				
		intSistLegado 		= Request("selSistLegado")					
		strRespTecLeg 		= Trim(Ucase(Request("txtRespTecLeg")))
		strRespFunLeg 		= Trim(Ucase(Request("txtRespFunLeg")))
		strTpDesligamento 	= Request("rdbTpDesligamento")
		strGerTecRespLeg 	= Trim(Ucase(Request("txtGerTecRespLeg")))

		'Response.write intPlano & "<br>"
		'Response.write intCDPlanoGeral & "<br>"
		'Response.write intSistLegado & "<br>"
		'Response.write strRespTecLeg & "<br>"
		'Response.write strRespFunLeg & "<br>"
		'Response.write strTpDesligamento & "<br>"
		'Response.write strGerTecRespLeg	& "<br>"			
		'Response.end															
										
		Call VerificaUsuarioExistente("Responsável Legado - Técnico", strRespTecLeg, "Legado", "PDS")			
		Call VerificaUsuarioExistente("Responsável Legado - Funcional", strRespFunLeg, "Legado", "PDS") 
															
		strSQL_AltPlanoPDS = ""
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & " UPDATE XPEP_PLANO_TAREFA_PDS SET"		
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & " SIST_NR_SEQUENCIAL_SISTEMA_LEGADO = " & intSistLegado 	
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", USUA_CD_USUARIO_RESP_LEG_TEC = '" & strRespTecLeg & "'"
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", USUA_CD_USUARIO_RESP_LEG_FUN = '" & strRespFunLeg & "'"
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", PPDS_TX_TIPO_DESLIGAMENTO ='" & strTpDesligamento & "'"
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", PPDS_TX_GER_TEC_RESP_LEGADO = '" & strGerTecRespLeg & "'"		
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPDS = strSQL_AltPlanoPDS & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 
			
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPDS)			
			
	end if

	if err.number = 0 then		
		strMSG = "Detalhamento alterado com sucesso."
		strGravado = 1
		'set correio = server.CreateObject("Persits.MailSender")
		'correio.host = "harpia.petrobras.com.br"
		 
		'correio.from="cursos@xproc.com"
	
		'correio.AddAddress "robson_28.infotec@petrobras.com.br"	     			
		'correio.AddAddress "gustavogomes.bearingpoint@petrobras.com.br"
		'correio.AddAddress "sergio.salomao.bearingpoint@petrobras.com.br"
						
		'correio.Subject="Inclusão de Novo Curso"
					
		'data_Atual=day(date) &"/"& month(date) &"/"& year(date)
						
		'correio.Body=" O Curso '" & UCASE(valor_cod) & "' foi INCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
		'correio.send
	else
		strMSG = "Houve um erro na alteração do detalhamento."
	end if	 
		
'************************************** EXCLUSÃO ************************************************	
elseif strAcao = "E" then	
	
	'*** PLANO DE COMUNICAÇÃO- EXCLUSÃO
	if strPlano = "PCM" then

		str_CdSeqPCM 		= request("pCdSeqPCM")		
													
		strSQL_ExcPlanoPCM = ""
		strSQL_ExcPlanoPCM = strSQL_ExcPlanoPCM & " DELETE XPEP_PLANO_TAREFA_PCM"		
		strSQL_ExcPlanoPCM = strSQL_ExcPlanoPCM & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano2
		strSQL_ExcPlanoPCM = strSQL_ExcPlanoPCM & " AND PPCM_NR_SEQUENCIA_TAREFA = " & str_CdSeqPCM										
		'Response.write strSQL_ExcPlanoPCM
		'Response.end
											
		on error resume next
			db_Cogest.Execute(strSQL_ExcPlanoPCM)		
				
	'*** PLANO DE AÇÕES CORRETIVAS E CONTINGÊNCIAS - EXCLUSÃO
	elseif strPlano = "PAC" then	
															
		intCdSeqPAC = Request("pCdSeqPAC")	
		strTipoExclusao = Request("pTipoExclusao")											
										
		'Response.write intCdSeqPAC & "<br>"
		'Response.write strTipoExclusao & "<br>"
		'Response.end									
															
		strSQL_ExcPlanoPAC = ""
		strSQL_ExcPlanoPAC = strSQL_ExcPlanoPAC & " DELETE XPEP_PLANO_TAREFA_PAC"		
		strSQL_ExcPlanoPAC = strSQL_ExcPlanoPAC & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC = strSQL_ExcPlanoPAC & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject							
		'Response.write strSQL_ExcPlanoPAC & "<br>"
		'Response.end	
		
		strSQL_ExcPlanoPAC_Sub = ""
		strSQL_ExcPlanoPAC_Sub = strSQL_ExcPlanoPAC_Sub & " DELETE XPEP_PLANO_TAREFA_PAC_FUNC"		
		strSQL_ExcPlanoPAC_Sub = strSQL_ExcPlanoPAC_Sub & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC_Sub = strSQL_ExcPlanoPAC_Sub & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		if strTipoExclusao = "Parcial" then
			strSQL_ExcPlanoPAC_Sub = strSQL_ExcPlanoPAC_Sub & " AND PPAC_NR_SEQUENCIA_FUNC = " & intCdSeqPAC			
		end if						
		'Response.write strSQL_ExcPlanoPAC_Sub & "<br>"
		'Response.end					
																
		on error resume next
		
			if strTipoExclusao = "Parcial" then 
				db_Cogest.Execute(strSQL_ExcPlanoPAC_Sub)
			elseif strTipoExclusao = "Total" then
				db_Cogest.Execute(strSQL_ExcPlanoPAC_Sub)
				db_Cogest.Execute(strSQL_ExcPlanoPAC)
			end if
			
	
	'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH - EXCLUSÃO
	elseif strPlano = "PAI" then		
		
		strSQL_Exc_Desenv_PAI = ""
		strSQL_Exc_Desenv_PAI = strSQL_Exc_Desenv_PAI & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"		
		strSQL_Exc_Desenv_PAI = strSQL_Exc_Desenv_PAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_Exc_Desenv_PAI = strSQL_Exc_Desenv_PAI & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		strSQL_Exc_Desenv_PAI = strSQL_Exc_Desenv_PAI & " AND PPCD_NR_SEQUENCIA_FUNC = '0'"
		'Response.write strSQL_Exc_Desenv_PAI & "<br><br>"
		'Response.end															
																						
		strSQL_ExcPlanoPAC_PAI = ""
		strSQL_ExcPlanoPAC_PAI = strSQL_ExcPlanoPAC_PAI & " DELETE XPEP_PLANO_TAREFA_PAC"		
		strSQL_ExcPlanoPAC_PAI = strSQL_ExcPlanoPAC_PAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC_PAI = strSQL_ExcPlanoPAC_PAI & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoPAC_PAI & "<br><br>"
		'Response.end				
					
		strSQL_ExcPlanoPCM_PAI = ""
		strSQL_ExcPlanoPCM_PAI = strSQL_ExcPlanoPCM_PAI & " DELETE XPEP_PLANO_TAREFA_PCM"		
		strSQL_ExcPlanoPCM_PAI = strSQL_ExcPlanoPCM_PAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		'Response.write strSQL_ExcPlanoPCM_PAI & "<br><br>"
		'Response.end						
					
		strSQL_ExcPlanoPAI = ""
		strSQL_ExcPlanoPAI = strSQL_ExcPlanoPAI & " DELETE XPEP_PLANO_TAREFA_PAI"		
		strSQL_ExcPlanoPAI = strSQL_ExcPlanoPAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAI = strSQL_ExcPlanoPAI & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		'Response.write strSQL_ExcPlanoPAI & "<br><br>"
		'Response.end
		
		strSQL_ExcPlanoGeral_PAI = ""
		strSQL_ExcPlanoGeral_PAI = strSQL_ExcPlanoGeral_PAI & " DELETE XPEP_PLANO_TAREFA_GERAL"		
		strSQL_ExcPlanoGeral_PAI = strSQL_ExcPlanoGeral_PAI & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoGeral_PAI = strSQL_ExcPlanoGeral_PAI & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoGeral_PAI & "<br><br>"
		'Response.end							
										
		on error resume next
			
			db_Cogest.Execute(strSQL_Exc_Desenv_PAI)	
			db_Cogest.Execute(strSQL_ExcPlanoPAC_PAI)	
			db_Cogest.Execute(strSQL_ExcPlanoPCM_PAI)	
			db_Cogest.Execute(strSQL_ExcPlanoPAI)	
			db_Cogest.Execute(strSQL_ExcPlanoGeral_PAI)
	
	'*** PLANO DE CONVERSÕES DE DADOS - EXCLUSÃO
	elseif strPlano = "PCD" then		
						
		'Response.write intDesenv & " - " & intPlano & " - " & intIdTaskProject & " - " & Request("pOndaRH")
		'Response.end
						
		strSQL_Exc_Desenv_PCD = ""
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"		
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		'if Request("pOndaRH") <> "RH" then
		'	strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " AND PPCD_NR_SEQUENCIA_FUNC =" & intDesenv
		'end if
		'Response.write strSQL_Exc_Desenv_PCD & "<br><br>"
		'Response.end															
																						
		strSQL_ExcPlanoPAC_PCD = ""
		strSQL_ExcPlanoPAC_PCD = strSQL_ExcPlanoPAC_PCD & " DELETE XPEP_PLANO_TAREFA_PAC"		
		strSQL_ExcPlanoPAC_PCD = strSQL_ExcPlanoPAC_PCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC_PCD = strSQL_ExcPlanoPAC_PCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoPAC_PCD & "<br><br>"
		'Response.end				
					
		strSQL_ExcPlanoPCM_PCD = ""
		strSQL_ExcPlanoPCM_PCD = strSQL_ExcPlanoPCM_PCD & " DELETE XPEP_PLANO_TAREFA_PCM"		
		strSQL_ExcPlanoPCM_PCD = strSQL_ExcPlanoPCM_PCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		'Response.write strSQL_ExcPlanoPCM_PCD & "<br><br>"
		'Response.end						
										
		strSQL_ExcPlanoPCD_Sub = ""
		strSQL_ExcPlanoPCD_Sub = strSQL_ExcPlanoPCD_Sub & " DELETE XPEP_PLANO_TAREFA_PCD_FUNC"		
		strSQL_ExcPlanoPCD_Sub = strSQL_ExcPlanoPCD_Sub & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPCD_Sub = strSQL_ExcPlanoPCD_Sub & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		'Response.write strSQL_ExcPlanoPCD & "<br><br>"
		'Response.end	
										
		strSQL_ExcPlanoPCD = ""
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " DELETE XPEP_PLANO_TAREFA_PCD"		
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		'Response.write strSQL_ExcPlanoPCD & "<br><br>"
		'Response.end
		
		strSQL_ExcPlanoGeral_PCD = ""
		strSQL_ExcPlanoGeral_PCD = strSQL_ExcPlanoGeral_PCD & " DELETE XPEP_PLANO_TAREFA_GERAL"		
		strSQL_ExcPlanoGeral_PCD = strSQL_ExcPlanoGeral_PCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoGeral_PCD = strSQL_ExcPlanoGeral_PCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoGeral_PCD & "<br><br>"
		'Response.end							
										
		on error resume next
			db_Cogest.Execute(strSQL_Exc_Desenv_PCD)
			db_Cogest.Execute(strSQL_ExcPlanoPAC_PCD)	
			db_Cogest.Execute(strSQL_ExcPlanoPCM_PCD)	
			db_Cogest.Execute(strSQL_ExcPlanoPCD_Sub)
			db_Cogest.Execute(strSQL_ExcPlanoPCD)	
			db_Cogest.Execute(strSQL_ExcPlanoGeral_PCD)
	
	'*** PLANO DE DESLIGAMENTO DE SISTEMAS LEGADOS - EXCLUSÃO
	elseif strPlano = "PDS" then		
		
		strSQL_Exc_Desenv_PDS = ""
		strSQL_Exc_Desenv_PDS = strSQL_Exc_Desenv_PDS & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"		
		strSQL_Exc_Desenv_PDS = strSQL_Exc_Desenv_PDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_Exc_Desenv_PDS = strSQL_Exc_Desenv_PDS & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		strSQL_Exc_Desenv_PDS = strSQL_Exc_Desenv_PDS & " AND PPCD_NR_SEQUENCIA_FUNC = '0'"	
		'Response.write strSQL_Exc_Desenv_PDS & "<br><br>"
		'Response.end															
																						
		strSQL_ExcPlanoPAC_PDS = ""
		strSQL_ExcPlanoPAC_PDS = strSQL_ExcPlanoPAC_PDS & " DELETE XPEP_PLANO_TAREFA_PAC"		
		strSQL_ExcPlanoPAC_PDS = strSQL_ExcPlanoPAC_PDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC_PDS = strSQL_ExcPlanoPAC_PDS & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoPAC_PDS & "<br><br>"
		'Response.end				
					
		strSQL_ExcPlanoPCM_PDS = ""
		strSQL_ExcPlanoPCM_PDS = strSQL_ExcPlanoPCM_PDS & " DELETE XPEP_PLANO_TAREFA_PCM"		
		strSQL_ExcPlanoPCM_PDS = strSQL_ExcPlanoPCM_PDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		'Response.write strSQL_ExcPlanoPCM_PDS & "<br><br>"
		'Response.end						
					
		strSQL_ExcPlanoPDS_Func = ""
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " DELETE XPEP_PLANO_TAREFA_PDS_FUNC"		
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject				
		'Response.write strSQL_ExcPlanoPDS_Func & "<br><br>"
		'Response.end
					
		strSQL_ExcPlanoPDS = ""
		strSQL_ExcPlanoPDS = strSQL_ExcPlanoPDS & " DELETE XPEP_PLANO_TAREFA_PDS"		
		strSQL_ExcPlanoPDS = strSQL_ExcPlanoPDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPDS = strSQL_ExcPlanoPDS & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		'Response.write strSQL_ExcPlanoPDS & "<br><br>"
		'Response.end
		
		strSQL_ExcPlanoGeral_PDS = ""
		strSQL_ExcPlanoGeral_PDS = strSQL_ExcPlanoGeral_PDS & " DELETE XPEP_PLANO_TAREFA_GERAL"		
		strSQL_ExcPlanoGeral_PDS = strSQL_ExcPlanoGeral_PDS & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoGeral_PDS = strSQL_ExcPlanoGeral_PDS & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoGeral_PDS & "<br><br>"
		'Response.end							
										
		on error resume next
			db_Cogest.Execute(strSQL_Exc_Desenv_PDS)
			db_Cogest.Execute(strSQL_ExcPlanoPAC_PDS)	
			db_Cogest.Execute(strSQL_ExcPlanoPCM_PDS)				
			db_Cogest.Execute(strSQL_ExcPlanoPDS_Func)
			db_Cogest.Execute(strSQL_ExcPlanoPDS)	
			db_Cogest.Execute(strSQL_ExcPlanoGeral_PDS)
					
	'*** PLANO DE PARADA OPERACIONAL - EXCLUSÃO
	elseif strPlano = "PPO" then		
		
		strSQL_Exc_Desenv_PPO = ""
		strSQL_Exc_Desenv_PPO = strSQL_Exc_Desenv_PPO & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"		
		strSQL_Exc_Desenv_PPO = strSQL_Exc_Desenv_PPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_Exc_Desenv_PPO = strSQL_Exc_Desenv_PPO & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		strSQL_Exc_Desenv_PPO = strSQL_Exc_Desenv_PPO & " AND PPCD_NR_SEQUENCIA_FUNC = '0'"	
		'Response.write strSQL_Exc_Desenv_PPO & "<br><br>"
		'Response.end															
																						
		strSQL_ExcPlanoPAC_PPO = ""
		strSQL_ExcPlanoPAC_PPO = strSQL_ExcPlanoPAC_PPO & " DELETE XPEP_PLANO_TAREFA_PAC"		
		strSQL_ExcPlanoPAC_PPO = strSQL_ExcPlanoPAC_PPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPAC_PPO = strSQL_ExcPlanoPAC_PPO & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoPAC_PPO & "<br><br>"
		'Response.end				
					
		strSQL_ExcPlanoPCM_PPO = ""
		strSQL_ExcPlanoPCM_PPO = strSQL_ExcPlanoPCM_PPO & " DELETE XPEP_PLANO_TAREFA_PCM"		
		strSQL_ExcPlanoPCM_PPO = strSQL_ExcPlanoPCM_PPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		'Response.write strSQL_ExcPlanoPCM_PPO & "<br><br>"
		'Response.end						
					
		strSQL_ExcPlanoPPO = ""
		strSQL_ExcPlanoPPO = strSQL_ExcPlanoPPO & " DELETE XPEP_PLANO_TAREFA_PPO"		
		strSQL_ExcPlanoPPO = strSQL_ExcPlanoPPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPPO = strSQL_ExcPlanoPPO & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		'Response.write strSQL_ExcPlanoPPO & "<br><br>"
		'Response.end
		
		strSQL_ExcPlanoGeral_PPO = ""
		strSQL_ExcPlanoGeral_PPO = strSQL_ExcPlanoGeral_PPO & " DELETE XPEP_PLANO_TAREFA_GERAL"		
		strSQL_ExcPlanoGeral_PPO = strSQL_ExcPlanoGeral_PPO & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoGeral_PPO = strSQL_ExcPlanoGeral_PPO & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		'Response.write strSQL_ExcPlanoGeral_PPO & "<br><br>"
		'Response.end							
										
		on error resume next
			db_Cogest.Execute(strSQL_Exc_Desenv_PPO)
			db_Cogest.Execute(strSQL_ExcPlanoPAC_PPO)	
			db_Cogest.Execute(strSQL_ExcPlanoPCM_PPO)	
			db_Cogest.Execute(strSQL_ExcPlanoPPO)	
			db_Cogest.Execute(strSQL_ExcPlanoGeral_PPO)
							
	end if
			
	if err.number = 0 then		
		strMSG = "Detalhamento excluído com sucesso."
	else
		strMSG = "Houve um erro na exclusão do detalhamento."
	end if		
end if

Public Function VerificaUsuarioExistente(strCampo, strChave, strTipoResponsavel, strPlano) 
		
	sql_VerUsuario= ""	
	'if strTipoResponsavel = "Sinergia" then 
		'sql_VerUsuario = sql_VerUsuario & " SELECT USUA_CD_USUARIO"		
		'sql_VerUsuario = sql_VerUsuario & " FROM USUARIO "
		'sql_VerUsuario = sql_VerUsuario & " WHERE USUA_CD_USUARIO = '" & strChave & "'"
		
		sql_VerUsuario = sql_VerUsuario & " SELECT USUA_TX_CD_USUARIO"		
		sql_VerUsuario = sql_VerUsuario & " FROM XPEP_EQUIPE_SINERGIA "
		sql_VerUsuario = sql_VerUsuario & " WHERE USUA_TX_CD_USUARIO = '" & strChave & "'"
		
	'elseif strTipoResponsavel = "Legado" then			
		'sql_VerUsuario = sql_VerUsuario & " SELECT USMA_CD_USUARIO"		
		'sql_VerUsuario = sql_VerUsuario & " FROM USUARIO_MAPEAMENTO "
		'sql_VerUsuario = sql_VerUsuario & " WHERE USMA_TX_MATRICULA <> 0"
		'sql_VerUsuario = sql_VerUsuario & " AND USMA_CD_USUARIO = '" & strChave & "'"
	'end if
	
	set rds_VerUsuario = db_Cogest.Execute(sql_VerUsuario)
	
	if rds_VerUsuario.eof then		
		if strTipoResponsavel = "Legado" then		
			'*** VERIFICA USUÁRIO DO LEGADO
			sql_VerUsuarioLegado = ""
			sql_VerUsuarioLegado = sql_VerUsuarioLegado & " SELECT USMA_CD_USUARIO"		
			sql_VerUsuarioLegado = sql_VerUsuarioLegado & " FROM USUARIO_MAPEAMENTO "
			sql_VerUsuarioLegado = sql_VerUsuarioLegado & " WHERE USMA_TX_MATRICULA <> 0"
			sql_VerUsuarioLegado = sql_VerUsuarioLegado & " AND USMA_CD_USUARIO = '" & strChave & "'"
		
			set rds_VerUsuarioLegado = db_Cogest.Execute(sql_VerUsuarioLegado)
		
			if rds_VerUsuarioLegado.eof then					
				strMsg = "Favor verificar a chave informada (" & strChave & "). No campo " & strCampo & "!"		
				rds_VerUsuarioLegado.close
				set rds_VerUsuarioLegado = nothing	
				rds_VerUsuario.close
				set rds_VerUsuario = nothing		
				Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pPlano=" & strPlano	
			end if	
			rds_VerUsuarioLegado.close
			set rds_VerUsuarioLegado = nothing	
		else
			strMsg = "Favor verificar a chave informada (" & strChave & "). No campo " & strCampo & "!"					
			rds_VerUsuario.close
			set rds_VerUsuario = nothing		
			Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pPlano=" & strPlano
		end if
	end if	
	rds_VerUsuario.close
	set rds_VerUsuario = nothing
End function

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">	

</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../indexA_xpep.asp"><img src="../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->   
	<table width="849" height="195" border="0" cellpadding="5" cellspacing="5">
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
				
		  <td width="117" height="29"></td>
				
		  <td width="53" height="29" valign="middle" align="left"></td>
				
		  <td height="29" valign="middle" align="left" colspan="2"> 
		  <%if err.number=0 then%>
		  <b><font face="Verdana" color="#330099" size="2"><%=strMSG%></font></b> 
		  </td>				
			  </tr>
		  <%else%>    
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  <b><font face="Verdana" size="2" color="#800000"><%=strMSG%> - <%=err.description%></font></b> 
		  </td>
			  </tr>
			  <%end if%>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="../../../indexA_xpep.asp">
			<img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
			  </tr>			  
			  <tr>				
				  <td width="117" height="1"></td>						
				  <td width="53" height="1" valign="middle" align="left"></td>						
				  <td height="1" valign="middle" align="left" width="32"> 
					<a href="seleciona_plano.asp">
					</a><a href="seleciona_plano.asp?selOnda=<%=int_CD_Onda%>&selFases=<%=str_Fase%>&selPlano=<%=Request("pPlanoOriginal")%>&selPlano2=<%=intPlano2%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
						
				  <td height="1" valign="middle" align="left" width="629"> 
					<font face="Verdana" color="#330099" size="2">Retornar - Seleção para Detalhamento das Atividades</font>
				  </td>
			  </tr>
		 <%if strPlano = "PDS" and strAcao <> "E" then%>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pds.asp?pAcao=A&pPlano=<%=intPlano%>&pTArefa=<%=int_SeqCDTarefa%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar para Tela de detalhamento do PDS</font> </td>
      </tr>
		<% end if %>
		 <%if strPlano = "PCM" then%>
			  <tr>
				  <td width="117" height="1"></td>				
				  <td width="53" height="1" valign="middle" align="left"></td>					
				  <td height="1" valign="middle" align="left" width="32"> 
					<a href="inclui_altera_plano_pcm.asp?pPlano2=<%=intPlano2%>&pOnda=<%=int_CD_Onda%>">
					<img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a>
				  </td>					
				  <td height="1" valign="middle" align="left" width="629"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela de detalhamento do PCM</font>
				  </td>
			  </tr>
			<%end if%>
			
			
			<%if strPlano = "PAC" then%>
			  <tr>
				  <td width="117" height="1"></td>				
				  <td width="53" height="1" valign="middle" align="left"></td>					
				  <td height="1" valign="middle" align="left" width="32"> 
					  <%if strTipoExclusao = "Total" then%>
						<a href="inclui_altera_plano_pac_1.asp?pAcao=I&pPlano=<%=intPlano%>&pTArefa=<%=intIdTaskProject%>&pFase=<%=str_Fase%>&pCdProjProject=<%=int_Cd_Projeto_Project%>&pOnda=<%=intOnda%>&pPlano_Origem=<%=strNomePlanoOrigem%>">
					  <%else%>
						<a href="inclui_altera_plano_pac_1.asp?pAcao=A&pPlano=<%=intPlano%>&pTArefa=<%=intIdTaskProject%>&pFase=<%=str_Fase%>&pCdProjProject=<%=int_Cd_Projeto_Project%>&pOnda=<%=intOnda%>&pPlano_Origem=<%=strNomePlanoOrigem%>">
					  <%end if%>	
					  <img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a>
				  </td>					
				  <td height="1" valign="middle" align="left" width="629"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela de detalhamento do PAC</font>
				  </td>
			  </tr>
			<%end if%>
			
			  <tr>
					
			  <td width="117" height="1"></td>
					
			  <td width="53" height="1" valign="middle" align="left"></td>
					
			  <td height="1" valign="middle" align="left" colspan="2"> 
			  </td>
			  </tr>
			</table>
  <table width="614" border="0">
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>		
  </table>
  
  	<%
	db_Cogest.close
	set db_Cogest = nothing
	%>

  <p>&nbsp;</p>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>

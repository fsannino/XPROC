<%
Response.Expires=0

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim rdsMaxPlano, strPlano, intCDPlanoGeral

strGravado = 0

strPlano 			= Request("pPlano")
intIdTaskProject	= Request("idTaskProject")
intPlano			= Request("pintPlano")
strNomeAtividade	= Request("pNomeAtividade")
strDtInicioAtiv 	= Formatdatetime(Request("pDtInicioAtiv"), 2)
strDtFimAtiv 		= Formatdatetime(Request("pDtFimAtiv"), 2)
strAcao				= Trim(Request("pAcao"))

strMSG =  ""
				
'************************************** INCLUSÃO ************************************************
if strAcao = "I" then

	if strPlano = "PCM" or strPlano = "PAC"then	
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
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ) Values( " & intPlano & "," & intCDPlanoGeral & ","
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & intIdTaskProject & ",'" & strNomeAtividade & "',"
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'" & strDtInicioAtiv & "','" & strDtFimAtiv & "',"
		strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'I', '" & Session("CdUsuario") & "', GETDATE())" 
								
		'*** PLANO DE PARADA OPERACIONAL - INCLUSÃO
		if strPlano = "PPO" then
					
			txtDescrParada 		= Trim(Ucase(Request("txtDescrParada")))
			strRespTecSinGeral 	= Request("selRespTecSinGeral")
			strRespTecLegGeral 	= Request("selRespTecLegGeral")
			strRespFunLegGeral 	= Request("selRespFunLegGeral")			
			intTempParada 		= Request("txtTempParada")
			strUnidadeMedida	= Request("selUnidMedida")			
			strProcedParada 	= Trim(Ucase(Request("txtProcedParada")))		
			strUsuarioGestor 	= Request("selUsuarioGestor")			
					
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
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & " ) Values( " & intPlano & ","
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & intCDPlanoGeral & ",'" & txtDescrParada & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & intTempParada & ",'" & strUnidadeMedida & "','" & strProcedParada & "','" & strDtParadaLegado & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'" & strDtIniR3 & "','" & strDtLimiteAprov & "','" & strRespTecSinGeral & "','" & strRespTecLegGeral & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'" & strRespFunLegGeral & "','" & strUsuarioGestor & "',"
			strSQL_NovoPlanoPPO = strSQL_NovoPlanoPPO & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
		
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPPO)
		
		'*** PLANO DE COMUNICAÇÃO - INCLUSÃO
		elseif strPlano = "PCM" then			
			
			strComunicacao		= Request("selComunicacao")			
			strOqueComunicar 	= Trim(Ucase(Request("txtOqueComunicar")))
			strAQuemComunicar 	= Trim(Ucase(Request("txtAQuemComunicar")))
			strUnidadeOrgao 	= Trim(Ucase(Request("txtUnidadeOrgao")))			
			strQuandoOcorre 	= Trim(Ucase(Request("txtQuandoOcorre")))
			strRespConteudo		= Trim(Ucase(Request("txtRespConteudo")))		
			strRespDivulg		= Trim(Ucase(Request("txtRespDivulg")))
			strAprovadorPB		= Trim(Ucase(Request("txtAprovadorPB")))	
											
			strDia = ""		
			strMes = ""
			strAno = ""			
			vetDtAprovacao = 	split(Request("txtDtAprovacao"),"/")	
			strDia = vetDtAprovacao(0)
			strMes = vetDtAprovacao(1)
			strAno = vetDtAprovacao(2)	
			strDtAprovacao = strMes & "/" & strDia & "/" & strAno 
			
			'Response.write strComunicacao & "<br>"
			'Response.write strOqueComunicar & "<br>"
			'Response.write strAQuemComunicar & "<br>"
			'Response.write strUnidadeOrgao 	& "<br>"
			'Response.write strQuandoOcorre 	& "<br>"
			'Response.write strRespConteudo	& "<br>"
			'Response.write strRespDivulg	& "<br>"
			'Response.write strAprovadorPB	& "<br>"
			'Response.write strDtAprovacao 	& "<br>"
																
			strSQL_NovoPlanoPCM = ""
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " INSERT INTO XPEP_PLANO_TAREFA_PCM ( "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_TP_COMUNICACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_O_QUE_COMUNICAR "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_PARA_QUEM "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_UNID_ORGAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_QUANDO_OCORRE "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_RESP_CONTEUDO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_RESP_DIVULGACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_APROVADOR_PB "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_DT_APROVACAO "				
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & " ) Values( " & intPlano & "," & intCDPlanoGeral & ",'" & strComunicacao & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'" & strOqueComunicar & "','" & strAQuemComunicar & "','" & strUnidadeOrgao & "','" & strQuandoOcorre & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'" & strRespConteudo & "','" & strRespDivulg & "','" & strAprovadorPB & "','" & strDtAprovacao & "',"
			strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
				
			'Response.WRITE strSQL_NovoPlanoPCM
			'Response.END
			
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
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
			str_RespTecLegGeral	= Request("selRespTecLegGeral")	
			str_RespTecSinGeral	= Request("selRespTecSinGeral")
								
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtInicio_Pai = 	split(Request("txtDtInicio_Pai"),"/")	
			strDia = vetDtInicio_Pai(0)
			strMes = vetDtInicio_Pai(1)
			strAno = vetDtInicio_Pai(2)	
			strDtInicio_Pai = strMes & "/" & strDia & "/" & strAno 
			
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
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & " ) Values( " & intPlano & "," & intCDPlanoGeral & ",'" & strCdInterface & "','" & strGrupo & "',"
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
		
			strAcoesCorrConting 	= Trim(Ucase(Request("txtAcoesCorrConting")))	
			strNomeInterface 		= Trim(Ucase(Request("txtNomeInterface")))			
			strUsuarioResponsavel 	= Request("selUsuarioResponsavel")
			strRespTecSinGeral		= Request("selRespTecSinGeral")			
			strRespFunSinGeral		= Request("selRespFunSinGeral")										
												
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTAprovacao_PAC = 	split(Request("txtDTAprovacao_PAC"),"/")	
			strDia = vetDTAprovacao_PAC(0)
			strMes = vetDTAprovacao_PAC(1)
			strAno = vetDTAprovacao_PAC(2)	
			strDTAprovacao_PAC = strMes & "/" & strDia & "/" & strAno 
			
			'Response.write strAcoesCorrConting & "<br>"		
			'Response.write strDTAprovacao_PAC	& "<br>"
			'Response.write strUsuarioResponsavel 	& "<br>"
			'Response.write strRespTecSinGeral 	& "<br>"
			'Response.write strRespFunSinGeral	& "<br>"			
			'Response.end
								
			strSQL_NovoPlanoPAC = ""
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " INSERT INTO XPEP_PLANO_TAREFA_PAC ( "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_TX_ACOES_CORR_CONT "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", PPAC_DT_APROVACAO "		
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_TRAT_PROC "		
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_TEC "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_FUN "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & " ) Values( " & intPlano & "," & intCDPlanoGeral & ",'" & strAcoesCorrConting & "',"
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & "'" & strDTAprovacao_PAC & "','" & strUsuarioResponsavel & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
			strSQL_NovoPlanoPAC = strSQL_NovoPlanoPAC & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
										
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPAC)
		
		'*** PLANO DE CONVERSÕES DE DADOS - INCLUSÃO
		elseif strPlano = "PCD" then		
					
			strRespTecLegGeral 	= Request("selRespTecLegGeral")
			strRespFunLegGeral 	= Request("selRespFunLegGeral")
			strRespTecSinGeral 	= Request("selRespTecSinGeral")
			strRespFunSinGeral 	= Request("selRespFunSinGeral")
			strDesenvAssociados	= Request("pSistemas")
			strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))		
			intSistLegado 		= Request("selSistLegado")
			strTipoCarga		= Request("selTipoCarga")
			strTipoDados 		= Request("selTipoDados")
			strCaractDado 		= Request("selCaractDado")
			intExtracao_PCD		= Request("txtExtracao_PCD")			
			intCarga_PCD	 	= Request("txtCarga_PCD")			
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
			vetDTCarga = 	split(Request("txtDTCarga_PCD"),"/")	
			strDia = vetDTCarga(0)
			strMes = vetDTCarga(1)
			strAno = vetDTCarga(2)	
			srtDTCarga_PCD = strMes & "/" & strDia & "/" & strAno 
						
			'Response.write strRespTecSinGeral & "<br>"
			'Response.write strRespTecLegGeral & "<br>"
			'Response.write strDesenvAssociados & "<br>"
			'Response.write strDadoMigrado	& "<br>"
			'Response.write intSistLegado & "<br>"
			'Response.write strTipoCarga	& "<br>"
			'Response.write strTipoDados & "<br>"
			'Response.write strCaractDado & "<br>"
			'Response.write intExtracao_PCD	& "<br>"
			'Response.write intCarga_PCD	 & "<br>"
			'Response.write strArqCarga	& "<br>"		
			'Response.write intVolume & "<br>"
			'Response.write strDependencias	& "<br>"	
			'Response.write strComoExecuta	& "<br>"		
			'Response.write strDTExtracao_PCD & "<br>"
			'Response.write srtDTCarga_PCD			
			'Response.end	
												
			strSQL_NovoPlanoPCD = ""
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & " INSERT INTO XPEP_PLANO_TAREFA_PCD ( "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & " PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PLTA_NR_SEQUENCIA_TAREFA "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", SIST_NR_SEQUENCIAL_SISTEMA_LEGADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_DADO_A_SER_MIGRADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_DESENV_ASSOCIADOS "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_TIPO_ATIVIDADE "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_TIPO_DADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_CARAC_DADO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
			'strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
			'strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_ARQ_CARGA "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_NR_VOLUME "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_TX_DEPENDENCIAS "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_DT_EXTRACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_DT_CARGA "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_tx_COMO_EXECUTA "
			'strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
			'strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_TEC "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_TX_OPERACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_CD_NR_USUARIO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", ATUA_DT_ATUALIZACAO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & " ) Values( " & intPlano & ","
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & intCDPlanoGeral & "," & intSistLegado & ","
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strDesenvAssociados & "','" & strDadoMigrado & "','" & strTipoCarga & "','" & strTipoDados & "','" & strCaractDado & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & intExtracao_PCD & "," & intCarga_PCD & ",'" & strArqCarga & "'," & intVolume & ","
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strDependencias & "','" & strDTExtracao_PCD & "','" & srtDTCarga_PCD & "','" & strComoExecuta & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'" & strRespTecLegGeral & "','" & strRespFunLegGeral & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
				
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPCD)		

		elseif strPlano = "PDS" then		

			intSistLegado 	= Request("selSistLegado")					
			strRespTecLeg 	= Request("selRespTecLeg")
			strRespFunLeg 	= Request("selRespFunLeg")
			strTpDesligamento 	= Request("rdbTpDesligamento")
			strGerTecRespLeg 	= Request("txtGerTecRespLeg")
						
			Response.write intPlano & "<br>"
			Response.write intCDPlanoGeral & "<br>"
			Response.write intSistLegado & "<br>"
			Response.write strRespTecLeg & "<br>"
			Response.write strRespFunLeg & "<br>"
			Response.write strTpDesligamento & "<br>"
			Response.write strGerTecRespLeg	& "<br>"			
			'Response.end	
												
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
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ) Values( " 
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " " & intPlano 
			strSQL_NovoPlanoPDS = strSQL_NovoPlanoPDS & " ," & intCDPlanoGeral
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
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	if strPlano = "PPO" then
				
		txtDescrParada 		= Trim(Ucase(Request("txtDescrParada")))			
		strRespTecSinGeral 	= Request("selRespTecSinGeral")
		strRespTecLegGeral 	= Request("selRespTecLegGeral")
		strRespFunLegGeral 	= Request("selRespFunLegGeral")			
		intTempParada 		= Request("txtTempParada")
		strUnidadeMedida	= Request("selUnidMedida")			
		strProcedParada 	= Trim(Ucase(Request("txtProcedParada")))			
		strUsuarioGestor 	= Request("selUsuarioGestor")			
				
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
								
		strComunicacao		= Request("selComunicacao")			
		strOqueComunicar 	= Trim(Ucase(Request("txtOqueComunicar")))
		strAQuemComunicar 	= Trim(Ucase(Request("txtAQuemComunicar")))
		strUnidadeOrgao 	= Trim(Ucase(Request("txtUnidadeOrgao")))			
		strQuandoOcorre 	= Trim(Ucase(Request("txtQuandoOcorre")))
		strRespConteudo		= Trim(Ucase(Request("txtRespConteudo")))		
		strRespDivulg		= Trim(Ucase(Request("txtRespDivulg")))
		strAprovadorPB		= Trim(Ucase(Request("txtAprovadorPB")))	
							
		strDia = ""		
		strMes = ""
		strAno = ""			
		vetDtAprovacao = 	split(Request("txtDtAprovacao"),"/")	
		strDia = vetDtAprovacao(0)
		strMes = vetDtAprovacao(1)
		strAno = vetDtAprovacao(2)	
		strDtAprovacao = strMes & "/" & strDia & "/" & strAno 
		
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
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " PPCM_TX_TP_COMUNICACAO = '" & strComunicacao & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_O_QUE_COMUNICAR = '" & strOqueComunicar & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_PARA_QUEM = '" & strAQuemComunicar & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_UNID_ORGAO = '" & strUnidadeOrgao & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_QUANDO_OCORRE = '" & strQuandoOcorre & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_RESP_CONTEUDO = '" & strRespConteudo & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_RESP_DIVULGACAO = '" & strRespDivulg & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_TX_APROVADOR_PB = '" & strAprovadorPB & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", PPCM_DT_APROVACAO = '" & strDtAprovacao & "'"	
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		'strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa  intIdTaskProject
		strSQL_AltPlanoPCM = strSQL_AltPlanoPCM & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		
		'Response.write strSQL_AltPlanoPCM
		'Response.end
		
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCM)	
	
	'*** PLANO DE AÇÕES CORRETIVAS E CONTINGÊNCIAS - ALTERAÇÃO
	elseif strPlano = "PAC" then		
	
		'****** PEGARÁ O COD DA TAREFA LÁ DO PROJECT
		intCDPlanoGeral = intIdTaskProject
	
		strAcoesCorrConting 	= Trim(Ucase(Request("txtAcoesCorrConting")))	
		strNomeInterface 		= Trim(Ucase(Request("txtNomeInterface")))	
		strUsuarioResponsavel 	= Request("selUsuarioResponsavel")
		strRespTecSinGeral		= Request("selRespTecSinGeral")			
		strRespFunSinGeral		= Request("selRespFunSinGeral")										
											
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDTAprovacao_PAC = 	split(Request("txtDTAprovacao_PAC"),"/")	
		strDia = vetDTAprovacao_PAC(0)
		strMes = vetDTAprovacao_PAC(1)
		strAno = vetDTAprovacao_PAC(2)	
		strDTAprovacao_PAC = strMes & "/" & strDia & "/" & strAno 
		
		'Response.write strAcoesCorrConting & "<br>"		
		'Response.write strDTAprovacao_PAC	& "<br>"
		'Response.write strUsuarioResponsavel 	& "<br>"
		'Response.write strRespTecSinGeral 	& "<br>"
		'Response.write strRespFunSinGeral	& "<br>"			
		'Response.end
								
		strSQL_AltPlanoPAC = ""
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " UPDATE XPEP_PLANO_TAREFA_PAC SET"	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " PPAC_TX_ACOES_CORR_CONT = '" & strAcoesCorrConting & "'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", PPAC_DT_APROVACAO = '" & strDTAprovacao_PAC & "'"		
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_TRAT_PROC = '" & strUsuarioResponsavel & "'"		
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_TEC = '" & strRespTecSinGeral & "'"	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_FUN = '" & strRespFunSinGeral & "'"	
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPAC = strSQL_AltPlanoPAC & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject							
									
		'Response.write strSQL_AltPlanoPAC
		'Response.end
		
		on error resume next			
			db_Cogest.Execute(strSQL_AltPlanoPAC)
	
	
	'*** PLANO DE CONVERSÕES DE DADOS - ALTERAÇÃO
	elseif strPlano = "PCD" then
								
		strRespTecLegGeral 	= Request("selRespTecLegGeral")
		strRespFunLegGeral 	= Request("selRespFunLegGeral")
		strRespTecSinGeral 	= Request("selRespTecSinGeral")
		strRespFunSinGeral 	= Request("selRespFunSinGeral")
		strDesenvAssociados	= Request("pSistemas")
		strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))			
		intSistLegado 		= Request("selSistLegado")
		strTipoCarga		= Request("selTipoCarga")
		strTipoDados 		= Request("selTipoDados")
		strCaractDado 		= Request("selCaractDado")
		intExtracao_PCD		= Request("txtExtracao_PCD")			
		intCarga_PCD	 	= Request("txtCarga_PCD")			
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
		vetDTCarga = 	split(Request("txtDTCarga_PCD"),"/")	
		strDia = vetDTCarga(0)
		strMes = vetDTCarga(1)
		strAno = vetDTCarga(2)	
		srtDTCarga_PCD = strMes & "/" & strDia & "/" & strAno 
						
		'Response.write strRespTecLegGeral & "<br>"
		'Response.write strRespFunLegGeral & "<br>"				
		'Response.write strRespTecSinGeral & "<br>"
		'Response.write strRespFunSinGeral & "<br>"
		'Response.write strDesenvAssociados & "<br>"
		'Response.write strDadoMigrado	& "<br>"
		'Response.write intSistLegado & "<br>"
		'Response.write strTipoCarga	& "<br>"
		'Response.write strTipoDados & "<br>"
		'Response.write strCaractDado & "<br>"
		'Response.write intExtracao_PCD	& "<br>"
		'Response.write intCarga_PCD	 & "<br>"
		'Response.write strArqCarga	& "<br>"		
		'Response.write intVolume & "<br>"
		'Response.write strDependencias	& "<br>"	
		'Response.write strComoExecuta	& "<br>"		
		'Response.write strDTExtracao_PCD & "<br>"
		'Response.write srtDTCarga_PCD	& "<br><br><br>"
		'Response.end													
													
		strSQL_AltPlanoPCD = ""
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " UPDATE XPEP_PLANO_TAREFA_PCD SET"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " SIST_NR_SEQUENCIAL_SISTEMA_LEGADO = " & intSistLegado
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_DADO_A_SER_MIGRADO = '" & strDadoMigrado & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_DESENV_ASSOCIADOS = '" & strDesenvAssociados & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_TIPO_ATIVIDADE = '" & strTipoCarga & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_TIPO_DADO = '" & strTipoDados & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_CARAC_DADO = '" & strCaractDado & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO = " & intExtracao_PCD		
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA = " & intCarga_PCD
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_ARQ_CARGA = '" & strArqCarga & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_VOLUME = " & intVolume
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_DEPENDENCIAS = '" & strDependencias & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_DT_EXTRACAO = '" & strDTExtracao_PCD & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_DT_CARGA = '" & srtDTCarga_PCD & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_TX_COMO_EXECUTA = '" & strComoExecuta & "'"
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
		'strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_TEC = '" & strRespTecLegGeral & "'"
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_FUN = '" & strRespFunLegGeral & "'"			
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_TEC = '" & strRespTecSinGeral & "'"	
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_FUN = '" & strRespFunSinGeral & "'"	
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 
		
		'Response.write strSQL_AltPlanoPCD
		'Response.end
		
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCD)	
			
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
		str_RespTecLegGeral	= Request("selRespTecLegGeral")	
		str_RespTecSinGeral	= Request("selRespTecSinGeral")
								
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtInicio_Pai = 	split(Request("txtDtInicio_Pai"),"/")	
		strDia = vetDtInicio_Pai(0)
		strMes = vetDtInicio_Pai(1)
		strAno = vetDtInicio_Pai(2)	
		strDtInicio_Pai = strMes & "/" & strDia & "/" & strAno 
		
		Response.write strCdInterface & "<br>"
		Response.write strGrupo & "<br>"
		Response.write strTipoBatch & "<br>"
		Response.write strNomeInterface 	& "<br>"
		Response.write strPgrmEnvolv 	& "<br>"
		Response.write strPreRequisitos	& "<br>"	
		Response.write strRestricoes	& "<br>"
		Response.write strDependencias 	& "<br>"
		Response.write strRespAciona 	& "<br>"
		Response.write strProcedimento 	& "<br>"
		Response.write strDtInicio_Pai 	& "<br>"
		Response.write str_RespTecLegGeral 	& "<br>"
		Response.write str_RespTecSinGeral 	& "<br>"	
		Response.end
													
		strSQL_AltPlanoPCD = ""
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " UPDATE XPEP_PLANO_TAREFA_PAI SET"	
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
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPPO = strSQL_AltPlanoPPO & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_SeqCDTarefa 
		
		'Response.write strSQL_AltPlanoPCD
		'Response.end
		
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCD)	

	'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH - ALTERAÇÃO
	elseif strPlano = "PDS" then								
				
		intSistLegado 	= Request("selSistLegado")					
		strRespTecLeg 	= Request("selRespTecLeg")
		strRespFunLeg 	= Request("selRespFunLeg")
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
end if

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
<link href="../../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">	

</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../../img/000003.gif" width="19" height="21"></td>
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
	    <td width="1" height="1" bgcolor="#003366"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../../indexA_xpep.asp"><img src="../../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->   
	<table width="849" height="150" border="0" cellpadding="5" cellspacing="5">
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
			</a><a href="../../../indexA_xpep.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar - Seleção para Detalhamento das Atividades</font></td>
			  </tr>
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
	<%				
	if (strAcao = "I" and strPlano <> "PAC" and strGravado = 1) or (strAcao = "I" and strPlano <> "PCM" and strGravado = 1) then
	%>	
		<tr>
		  <td width="2">&nbsp;</td>
		  <td width="271"><div align="right" class="campob">Crie o plano de Conting&ecirc;ncia:</div>	      </td>
		  <td width="45"><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
		  <td width="235" valign="top" class="campob"><div align="right">Crie o plano de Comunica&ccedil;&atilde;o:</div></td>
		  <td width="39"><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
		</tr>
	<%
	end if
	%>
  </table>
  <p>&nbsp;</p>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>

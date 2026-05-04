<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim rdsMaxPlano, strPlano, intCDPlanoGeral

strPlano 			= Request("pPlano")
intIdTaskProject	= Request("idTaskProject")
intPlano			= Request("pintPlano")
strNomeAtividade	= Request("pNomeAtividade")
strDtInicioAtiv 	= Formatdatetime(Request("pDtInicioAtiv"), 2)
strDtFimAtiv 		= Formatdatetime(Request("pDtFimAtiv"), 2)
strAcao				= Trim(Request("pAcao"))

strMSG =  ""

if strAcao = "I" then

	strVerificaAtvidade = ""
	strVerificaAtvidade = strVerificaAtvidade & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
	strVerificaAtvidade = strVerificaAtvidade & " FROM XPEP_PLANO_TAREFA_GERAL"
	strVerificaAtvidade = strVerificaAtvidade & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & intIdTaskProject
	Set rdsVerificaAtvidade = db_Cogest.Execute(strVerificaAtvidade)			
	
	if not rdsVerificaAtvidade.EOF then
		strMSG = "Já existe detalhamento cadastrado para esta atividade."
	else	
		'*** Seleciona o cod para a Nova Tarefa Geral - na tabela XPEP_PLANO_TAREFA_GERAL		
		intCDPlanoGeral = 0	
		Set rdsMaxPlano = db_Cogest.Execute("SELECT MAX(PLTA_NR_SEQUENCIA_TAREFA) AS INT_MAIOR_TAREFA_GERAL FROM XPEP_PLANO_TAREFA_GERAL")			
		if isnull(rdsMaxPlano("INT_MAIOR_TAREFA_GERAL")) then
			intCDPlanoGeral = 1
		else
			intCDPlanoGeral = rdsMaxPlano("INT_MAIOR_TAREFA_GERAL") + 1
		end if
		rdsMaxPlano.Close	
		set rdsMaxPlano = nothing
		
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
		
		'Response.write strSQL_NovoPlanoGeral
		'Response.end
		
		'*** PLANO DE PARADA OPERACIONAL
		if strPlano = "PPO" then
		
			'intIdTaskProject	= Request("idTaskProject")
			'intPlano			= Request("pintPlano")
			txtDescrParada 		= Request("txtDescrParada")			
			strRespTecSinGeral 	= Request("selRespTecSinGeral")
			strRespTecLegGeral 	= Request("selRespTecLegGeral")
			strRespFunLegGeral 	= Request("selRespFunLegGeral")			
			intTempParada 		= Request("txtTempParada")
			strUnidadeMedida	= Request("selUnidMedida")			
			strProcedParada 	= Request("txtProcedParada")			
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
		
		'*** PLANO DE COMUNICAÇÃO
		elseif strPlano = "PCM" then	
				
			'intIdTaskProject	= Request("idTaskProject")
			'intPlano			= Request("pintPlano")
			strComunicacao		= Request("selComunicacao")			
			strOqueComunicar 	= Request("txtOqueComunicar")
			strAQuemComunicar 	= Request("txtAQuemComunicar")
			strUnidadeOrgao 	= Request("txtUnidadeOrgao")			
			strQuandoOcorre 	= Request("txtQuandoOcorre")
			strRespConteudo		= Request("txtRespConteudo")			
			strRespDivulg		= Request("txtRespDivulg")			
			strAprovadorPB		= Request("txtAprovadorPB")	
											
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
			'strSQL_NovoPlanoPCM = strSQL_NovoPlanoPCM & ", PPCM_TX_DESC_ATIVIDADE "
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
		
		'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH
		elseif strPlano = "PAI" then		
		
			strCdInterface		= Request("txtCdInterface")			
			strGrupo 			= Request("txtGrupo")
			strTipoBatch 		= Request("selTipoBatch")
			strNomeInterface 	= Request("txtNomeInterface")			
			strPgrmEnvolv 		= Request("txtPgrmEnvolv")
			strPreRequisitos	= Request("txtPreRequisitos")			
			strRestricoes		= Request("txtRestricoes")			
			strDependencias		= Request("txtDependencias")	
			strRespAciona		= Request("txtRespAciona")	
			strProcedimento		= Request("txtProcedimento")	
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
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'" & str_RespTecLegGeral & "','" & str_RespTecSinGeral & "',"	
			strSQL_NovoPlanoPAI = strSQL_NovoPlanoPAI & "'I','" & Session("CdUsuario") & "',GETDATE())" 		
									
			'Response.write strSQL_NovoPlanoPAI
			'Response.end
			
			on error resume next
				db_Cogest.Execute(strSQL_NovoPlanoGeral)
				db_Cogest.Execute(strSQL_NovoPlanoPAI)
				
		'*** PLANO DE ACIONAMENTO DE INTERFACES E PROCESSOS BATCH
		elseif strPlano = "PAC" then		
		
			strAcoesCorrConting 	= Request("txtAcoesCorrConting")		
			strNomeInterface 		= Request("txtNomeInterface")			
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
									
			'Response.write strSQL_NovoPlanoPAC
			'Response.end		
			
			on error resume next
				'db_Cogest.Execute(strSQL_NovoPlanoGeral)
				'db_Cogest.Execute(strSQL_NovoPlanoPAC)
		
		end if
	
		if err.number = 0 then
		
			strMSG = "Detalhamento cadastrado com sucesso."
		
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
	
elseif strAcao = "A" then

' ****************  ALTERAÇÃO *********************

end if
%>
<html>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBegin template="/Templates/BASICO_XPEP_01.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
	<!-- InstanceBeginEditable name="doctitle" -->
	<title>SINERGIA # XPROC # Processos de Negócio</title>
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">	
	<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaSci" -->		
<!-- InstanceEndEditable -->
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../../indexA.asp"><img src="../../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      
    <td colspan="3" height="20"><!-- InstanceBeginEditable name="Botao_01" -->
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:confirma_ppo()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>     
<!-- InstanceBeginEditable name="Corpo_Princ" -->
  <table border="0" width="849" height="81">
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
			<img border="0" src="../../../imagens/selecao_F02.gif" align="right"></a></td>
				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="seleciona_plano.asp">
			<img border="0" src="../../../imagens/selecao_F02.gif" align="right"></a></td>
				
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
  <table width="70%" border="0">
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
	if (strAcao = "I" and strPlano <> "PAC") and (strAcao = "I" and strPlano <> "PCM") then
	%>	
		<tr>
		  <td width="2%">&nbsp;</td>
		  <td width="40%"><div align="right" class="campob">Crie o plano de Conting&ecirc;ncia:</div></td>
		  <td width="5%"><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
		  <td width="45%" class="campob"><div align="right">Crie o plano de Comunica&ccedil;&atilde;o:</div></td>
		  <td width="8%"><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
		</tr>
	<%
	end if
	%>
  </table>
  <p>&nbsp;</p>
<!-- InstanceEndEditable -->
</body>

<!-- InstanceEnd --></html>
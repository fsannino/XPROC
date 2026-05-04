<%
Response.Expires=0

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim rdsMaxPlano, strPlano, intCDPlanoGeral

strGravado = 0

strAcao				= Trim(Request("pAcao"))
if strAcao <> "" then
	strPlano 			= Request("pPlano") 
	intPlano			= Request("pintPlano")  
	intIdTaskProject	= Request("idTaskProject")
	intCdSeqFunc   	    = Request("pCdSeqFunc")
else
	strAcao				= Trim(Request("pAcao2"))
	strPlano 			= Request("pPlano2") 
	intPlano			= Request("pintPlano2") 
	intIdTaskProject	= Request("idTaskProject2")
	intCdSeqFunc   	    = Request("pCdSeqFunc2")
end if
strNomeAtividade	= Request("pNomeAtividade")
strDtInicioAtiv 	= Formatdatetime(Request("pDtInicioAtiv"), 2)
strDtFimAtiv 		= Formatdatetime(Request("pDtFimAtiv"), 2)

strMSG =  ""

int_Cd_Projeto_Project 	= Request("pCdProjProject")
str_Cd_Onda 			= Request("pOnda")
str_Cd_Plano 			= Request("pPlano")
str_Fase 				= Request("pFase")

'response.Write("str_Atividade : " & str_Atividade & "<p>")
'response.Write("Ds_Plano : " & strPlano & "<p>")
'response.Write("Nr_Palno : " & intPlano & "<p>")
'response.Write("str_Cd_Plano : " & str_Cd_Plano & "<p>")
'response.Write("str_Fase : " & str_Fase & "<p>")
'response.Write("Ação : " & strAcao & "<p>")
'response.Write("Ds_Plano : " & strPlano & "<p>")
'response.Write("Nr_Palno : " & intPlano & "<p>")
'response.Write("Nr_Tarefa : " & intIdTaskProject & "<p>")
'response.Write("intIdTaskProject : " & intIdTaskProject & "<p>")
'Response.end

'************************************** INCLUSÃO ************************************************
if strAcao = "I" then

	blnNaoCadastraPlano = False
											
	'*** PLANO DE PARADA OPERACIONAL - INCLUSÃO
	if strPlano = "PDS" then
				
		str_FuncDesat = Request("txtFuncDesat")
		dat_DtDesliga = Request("txtDtDesliga")
		hor_HrDesliga = Right("00" & (Request("txtHrDesliga")),2)
		hor_MnDesliga = Right("00" & (Request("txtmnDesliga")),2)
		str_ProcDesl = Request("txtProcDesl")
		str_DestDados = Request("txtDestDados")

		'Response.write "aaa = FuncDesat=" & str_FuncDesat & "<br>"
		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.write "HrDesliga" & hor_HrDesliga & "<br>"
		'Response.write "MnDesliga" & hor_MnDesliga & "<br>"
		'Response.write "ProcDesl" & str_ProcDesl & "<br>"
		'Response.write "DestDados" & str_DestDados & "<br>"
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataParada = 	split(dat_DtDesliga,"/")	
		strDia = vetDtDataParada(0)
		strMes = vetDtDataParada(1)
		strAno = vetDtDataParada(2)	
		dat_DtDesliga = strMes & "/" & strDia & "/" & strAno 					

		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.end	

		'*** Seleciona o cod para a Nova Seq para Sub-Atividade de PDS
		intCdSeqFunc = 0	
		str_SQL = ""
		str_SQL = str_SQL & " SELECT "
		str_SQL = str_SQL & " MAX(PPDS_NR_SEQUENCIA_FUNC) AS int_Max_SeqFunc "
		str_SQL = str_SQL & " FROM XPEP_PLANO_TAREFA_PDS_FUNC "
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
													
		strSQL_Nova_Funcionalidade = ""
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " INSERT INTO XPEP_PLANO_TAREFA_PDS_FUNC ( "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " PLAN_NR_SEQUENCIA_PLANO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PLTA_NR_SEQUENCIA_TAREFA "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_NR_SEQUENCIA_FUNC "		
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_FUNC_DESATIVADAS "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_DT_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_HR_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_PROC_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_DEST_DD_TEMPO_RETENCAO "			
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_TX_OPERACAO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_CD_NR_USUARIO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_DT_ATUALIZACAO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " ) Values( " 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & intPlano 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", " & intIdTaskProject 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", " & intCdSeqFunc
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_FuncDesat) & "'"
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & dat_DtDesliga & "'"
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '"  & hor_HrDesliga & ":" & hor_MnDesliga & "'" 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_ProcDesl) & "'" 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_DestDados) & "'"			
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", 'I','" & Session("CdUsuario") & "',GETDATE())" 		
	
		'Response.write strSQL_Nova_Funcionalidade
		'Response.end
	
		on error resume next
			db_Cogest.Execute(strSQL_Nova_Funcionalidade)
	
	'*** PLANO DE COMUNICAÇÃO - INCLUSÃO
	elseif strPlano = "PCM" then			
	
	'*** PLANO DE CONVERSÕES DE DADOS - INCLUSÃO
	elseif strPlano = "PCD" then		
	
		strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
		strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))
		strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
		strRespFunSinGeral 	= Trim(Ucase(Request("txtRespFunSinGeral")))
		
		strDesenvAssociados	= Request("pSistemas")
		strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))					
		strSistLegado 		= Trim(Ucase(Request("txtSistLegado")))			
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
		
		str_sqlVerificaPlano = ""
		str_sqlVerificaPlano = str_sqlVerificaPlano & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
		str_sqlVerificaPlano = str_sqlVerificaPlano & ", PLTA_NR_SEQUENCIA_TAREFA "					
		str_sqlVerificaPlano = str_sqlVerificaPlano & " FROM XPEP_PLANO_TAREFA_GERAL"
		str_sqlVerificaPlano = str_sqlVerificaPlano & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		str_sqlVerificaPlano = str_sqlVerificaPlano & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		'Response.write str_sqlVerificaPlano & "<br>"
		'Response.end 		
		set rdsVerificaPlano = db_Cogest.Execute(str_sqlVerificaPlano)	
		
		'*** CASO NÃO EXISTA REGISTRO
		if rdsVerificaPlano.eof then					
			'*** INCLUSÃO DO PLANO - TAREFA GERAL (XPEP_PLANO_TAREFA_GERAL)
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
			strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & " ) Values(" & intPlano & "," & intIdTaskProject & ","
			strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & intIdTaskProject & ",'" & strNomeAtividade & "',"
			strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'" & strDtInicioAtiv & "','" & strDtFimAtiv & "',"
			strSQL_NovoPlanoGeral = strSQL_NovoPlanoGeral & "'I', '" & Session("CdUsuario") & "', GETDATE())" 		
			'Response.write strSQL_NovoPlanoGeral
			'Response.end
			db_Cogest.Execute(strSQL_NovoPlanoGeral)		
		end if				
		rdsVerificaPlano.close
		set rdsVerificaPlano  = nothing
		
		'*** VERIFICA SE O PLANO JÁ FOI CADASTRADO EM XPEP_PLANO_TAREFA_PCD
		str_sqlAtividade = ""
		str_sqlAtividade = str_sqlAtividade & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
		str_sqlAtividade = str_sqlAtividade & ", PLTA_NR_SEQUENCIA_TAREFA "					
		str_sqlAtividade = str_sqlAtividade & " FROM XPEP_PLANO_TAREFA_PCD"
		str_sqlAtividade = str_sqlAtividade & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		str_sqlAtividade = str_sqlAtividade & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		'Response.write str_sqlAtividade
		'Response.end
		set rdsVerificaAtividadePCD = db_Cogest.Execute(str_sqlAtividade)	
		
		if rdsVerificaAtividadePCD.eof then
			'*** GRAVA O PCD NA TABELA XPEP_PLANO_TAREFA_PCD
			strSQL_NovoPlanoPCD = ""
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "INSERT INTO XPEP_PLANO_TAREFA_PCD ("
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & "PLAN_NR_SEQUENCIA_PLANO "
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & ", PLTA_NR_SEQUENCIA_TAREFA)"
			strSQL_NovoPlanoPCD = strSQL_NovoPlanoPCD & " Values(" & intPlano & "," & intIdTaskProject & ")"			
			'Response.write strSQL_NovoPlanoPCD
			'Response.end 
			db_Cogest.Execute(strSQL_NovoPlanoPCD)
		end if
		rdsVerificaAtividadePCD.close
		set rdsVerificaAtividadePCD = nothing
		
		'*** Seleciona o cod para a Nova Seq para Sub-Atividade de PCD
		intCdSeqFunc = 0	
		str_SQL = ""
		str_SQL = str_SQL & " SELECT "
		str_SQL = str_SQL & " MAX(PPCD_NR_SEQUENCIA_FUNC) AS int_Max_SeqFunc "
		str_SQL = str_SQL & " FROM XPEP_PLANO_TAREFA_PCD_FUNC "
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
		
		'Response.write intCdSeqFunc & "<br>"	
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
											
		strSQL_NovoPlanoPCD_Sub = ""
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & " INSERT INTO XPEP_PLANO_TAREFA_PCD_FUNC ("
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & "PLAN_NR_SEQUENCIA_PLANO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PLTA_NR_SEQUENCIA_TAREFA "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_NR_SEQUENCIA_FUNC "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_SISTEMA_LEGADO "		
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_DADO_A_SER_MIGRADO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_TIPO_ATIVIDADE "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_TIPO_DADO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_CARAC_DADO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_ARQ_CARGA "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_NR_VOLUME "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_TX_DEPENDENCIAS "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_DT_EXTRACAO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_DT_CARGA_INICIO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_DT_CARGA_FIM "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", PPCD_tx_COMO_EXECUTA "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", USUA_CD_USUARIO_RESP_LEG_TEC "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", ATUA_TX_OPERACAO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", ATUA_CD_NR_USUARIO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ", ATUA_DT_ATUALIZACAO "
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & ") Values(" & intPlano & ","
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & intIdTaskProject & "," & intCdSeqFunc & ",'" & strSistLegado & "',"
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & "'" & strDadoMigrado & "','" & strTipoCarga & "','" & strTipoDados & "','" & strCaractDado & "',"
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & intExtracao_PCD & ",'" & strExtracao_Unid & "'," & intCarga_PCD & ",'" & strCarga_Unid & "','" & strArqCarga & "'," & intVolume & ","
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & "'" & strDependencias & "','" & strDTExtracao_PCD & "','" & srtDTCarga_PCD_Ini & "','" & srtDTCarga_PCD_Fim & "','" & strComoExecuta & "',"
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & "'" & strRespTecLegGeral & "','" & strRespFunLegGeral & "','" & strRespTecSinGeral & "','" & strRespFunSinGeral & "',"
		strSQL_NovoPlanoPCD_Sub = strSQL_NovoPlanoPCD_Sub & "'I','" & Session("CdUsuario") & "',GETDATE())" 						
		'Response.write strSQL_NovoPlanoPCD_Sub		
		'Response.end				
		on error resume next			
			db_Cogest.Execute(strSQL_NovoPlanoPCD_Sub)	
			
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
						sql_NovoDesenvAss = sql_NovoDesenvAss & intIdTaskProject & "," & intCdSeqFunc & ",'" & vetDesenvAssociados(i) & "',"
						sql_NovoDesenvAss = sql_NovoDesenvAss & "'I','" & Session("CdUsuario") & "',GETDATE())" 
						'Response.write strSQL_NovoPlanoPCD_Sub		
						'Response.end			
						db_Cogest.Execute(sql_NovoDesenvAss)						
					end if					
				next
			end if			
	end if

	if err.number = 0 then
	
		strMSG = "Registro incluido com sucesso."
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
		strMSG = "Houve um erro no cadastro do registro."
	end if	 

'************************************** ALTERAÇÃO ************************************************	
elseif strAcao = "A" then
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	if strPlano = "PDS" then
				
		str_FuncDesat = Request("txtFuncDesat")
		dat_DtDesliga = Request("txtDtDesliga")
		hor_HrDesliga = Right("00" & (Request("txtHrDesliga")),2)
		hor_MnDesliga = Right("00" & (Request("txtmnDesliga")),2)
		str_ProcDesl = Request("txtProcDesl")
		str_DestDados = Request("txtDestDados")

		'Response.write "FuncDesat=" & str_FuncDesat & "<br>"
		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.write "HrDesliga" & hor_HrDesliga & "<br>"
		'Response.write "MnDesliga" & hor_MnDesliga & "<br>"
		'Response.write "ProcDesl" & str_ProcDesl & "<br>"
		'Response.write "DestDados" & str_DestDados & "<br>"
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataParada = 	split(dat_DtDesliga,"/")	
		strDia = vetDtDataParada(0)
		strMes = vetDtDataParada(1)
		strAno = vetDtDataParada(2)	
		dat_DtDesliga = strMes & "/" & strDia & "/" & strAno 					
													
		strSQL_AltPlanoPDS_Func = ""
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " UPDATE XPEP_PLANO_TAREFA_PDS_FUNC SET"		
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " PPDS_TX_FUNC_DESATIVADAS = '" & UCase(str_FuncDesat) & "'"	
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_DT_DESLIGAMENTO = '" & dat_DtDesliga & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_HR_DESLIGAMENTO = '" & hor_HrDesliga & ":" & hor_MnDesliga & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_PROC_DESLIGAMENTO ='" & UCase(str_ProcDesl) & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_DEST_DD_TEMPO_RETENCAO = '" & UCase(str_DestDados) & "'"		
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject 
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " AND PPDS_NR_SEQUENCIA_FUNC = " & intCdSeqFunc 
		'response.Write(strSQL_AltPlanoPDS_Func)
		'Response.end
			
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPDS_Func)			
		
	'*** PLANO DE COMUNICAÇÃO- ALTERAÇÃO
	elseif strPlano = "PCM" then								
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	elseif strPlano = "PCD" then		
			
		strRespTecLegGeral 	= Trim(Ucase(Request("txtRespTecLegGeral")))
		strRespFunLegGeral 	= Trim(Ucase(Request("txtRespFunLegGeral")))
		strRespTecSinGeral 	= Trim(Ucase(Request("txtRespTecSinGeral")))
		strRespFunSinGeral 	= Trim(Ucase(Request("txtRespFunSinGeral")))
		
		strDesenvAssociados	= Request("pSistemas")
		strDadoMigrado		= Trim(Ucase(Request("txtDadoMigrado")))					
		strSistLegado 		= Trim(Ucase(Request("txtSistLegado")))			
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
					
		intCdSeqPCD 		= Request("pCdSeqPCD")			
					
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
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " UPDATE XPEP_PLANO_TAREFA_PCD_FUNC SET"	
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
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		strSQL_AltPlanoPCD = strSQL_AltPlanoPCD & " AND PPCD_NR_SEQUENCIA_FUNC = " & intCdSeqPCD 			
		'Response.write strSQL_AltPlanoPCD
		'Response.end	
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPCD)	
			
			'*** EXCLUSÃO DOS DESENVOLVIMENTOS ASSOCIADOS PARA ESTE PLANO ***				
			sql_DelDesenvAss = ""
			sql_DelDesenvAss = sql_DelDesenvAss & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"
			sql_DelDesenvAss = sql_DelDesenvAss & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano 
			sql_DelDesenvAss = sql_DelDesenvAss & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject		
			sql_DelDesenvAss = sql_DelDesenvAss & " AND PPCD_NR_SEQUENCIA_FUNC = " & intCdSeqPCD	
			
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
						sql_NovoDesenvAss = sql_NovoDesenvAss & intIdTaskProject & "," & intCdSeqPCD & ",'" & vetDesenvAssociados(i) & "',"
						sql_NovoDesenvAss = sql_NovoDesenvAss & "'I','" & Session("CdUsuario") & "',GETDATE())"
						db_Cogest.Execute(sql_NovoDesenvAss)						
					end if					
				next
			end if				
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


elseif strAcao = "E" then
			
	'*** PLANO DE  - EXCLUSÃO
	if strPlano = "PDS" then
													
		strSQL_ExcPlanoPDS_Func = ""
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " DELETE FROM XPEP_PLANO_TAREFA_PDS_FUNC "		
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject 
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " AND PPDS_NR_SEQUENCIA_FUNC = " & intCdSeqFunc 
		'response.Write(strSQL_ExcPlanoPDS_Func)
		'Response.end
			
		on error resume next
			db_Cogest.Execute(strSQL_ExcPlanoPDS_Func)			
		
	'*** PLANO DE COMUNICAÇÃO- EXCLUSÃO
	elseif strPlano = "PCM" then								
			
	'*** PLANO DE CONVERSÕES DE DADOS - EXCLUSÃO
	elseif strPlano = "PCD" then	
	
		intCdSeqPCD 		  = Request("pDesenv")
	
		strSQL_Exc_Desenv_PCD = ""
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " DELETE XPEP_TAREFA_DESENVOLVIMENTO"		
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject	
		strSQL_Exc_Desenv_PCD = strSQL_Exc_Desenv_PCD & " AND PPCD_NR_SEQUENCIA_FUNC =" & intCdSeqPCD		
		'Response.write strSQL_Exc_Desenv_PCD & "<br><br>"
		'Response.end	
	
		strSQL_ExcPlanoPCD = ""
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " DELETE XPEP_PLANO_TAREFA_PCD_FUNC"		
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject										
		strSQL_ExcPlanoPCD = strSQL_ExcPlanoPCD & " AND PPCD_NR_SEQUENCIA_FUNC = " & intCdSeqPCD
		'Response.write strSQL_ExcPlanoPCD & "<br><br>"
		'Response.end	
			
		on error resume next
			db_Cogest.Execute(strSQL_Exc_Desenv_PCD)	
			db_Cogest.Execute(strSQL_ExcPlanoPCD)		
	end if

	if err.number = 0 then		
		strMSG = "Detalhamento excluído com sucesso."
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
	<table width="849" height="207" border="0" cellpadding="5" cellspacing="5">
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
			  <% if strPlano = "PDS" then 
			        if str_Acao = "I" then %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pds_func.asp?pAcao=I&pPlano=<%=intPlano%>&pIdTaskProject=<%=intIdTaskProject%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"> <font face="Verdana" color="#330099" size="2">Retornar - Cadastro de mais uma Funcionalidade</font></td>
	  </tr>
	  				<% end if %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pds.asp?pAcao=A&pPlano=<%=intPlano%>&pTArefa=<%=intIdTaskProject%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar - Detalhamento de PDS </font></td>
      </tr>
	  
	  <% elseif strPlano = "PCD" then%> 
		 	<tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pcd_rh.asp?pAcao=A&pPlano=<%=intPlano%>&pTArefa=<%=intIdTaskProject%>&pFase=<%=str_Fase%>&pCdProjProject=<%=int_Cd_Projeto_Project%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar - Detalhamento de PCD </font></td>   				 	
      		</tr>	  
	  <% else %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="seleciona_plano.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar - Cadastramento de mais um PCM para este Plano </font></td>
      </tr>
	  <% end if %>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="seleciona_plano.asp">
			</a><a href="seleciona_plano.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				
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
	if (strAcao = "I" and strPlano <> "PAC" and strGravado = 1) and (strAcao = "I" and strPlano <> "PCM" and strGravado = 1) then
	%>	
		<tr>
		  <td width="2">&nbsp;</td>
		  <td width="271">&nbsp;</td>
		  <td width="45">&nbsp;</td>
		  <td width="235" valign="top" class="campob">&nbsp;</td>
		  <td width="39">&nbsp;</td>
		</tr>
	<%
	end if
	%>
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

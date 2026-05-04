<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strMostraUnidade = "Não"
		
if trim(Request("parAcao")) <> "" then
	strAcao 	=  trim(Request("parAcao"))
elseif trim(Request("pAcao")) <> "" then
	strAcao 	=  trim(Request("pAcao"))
end if

strGrava		=  trim(Request("parGrava"))

if trim(Request("hdSala")) <> "" then
	strSala 	=  trim(Request("hdSala"))
elseif trim(Request("pSala")) <> "" then
	strSala 	=  trim(Request("pSala"))
end if

if Request("selCorte") <> "" then
	session("Corte")=  trim(Request("selCorte"))	
end if

intCdCorte	=  session("Corte")

if Request("selCurso") <> "0" and Request("selCurso") <> "" Then
	strCurso = Request("selCurso")
	
	'*** ====== VERIFICA O TOTAL DE PERÍODOS DO CURSO =========================================
	set rsPeriodoCurso = db_banco.execute("SELECT CURS_NUM_CARGA_CURSO FROM GRADE_CURSO WHERE CURS_CD_CURSO='" & strCurso & "' AND CORT_CD_CORTE = " & intCdCorte)
	strCarga = rsPeriodoCurso("CURS_NUM_CARGA_CURSO")		
	strPeriodos = strCarga / 4
	strDias = strPeriodos / 2
	
	rsPeriodoCurso.close
	set rsPeriodoCurso = nothing
else
	strCurso = 0
end if

strMultiplic = 0
if trim(Request("selMultiplic")) <> "" Then
	strMultiplic = trim(Request("selMultiplic"))
elseif trim(Request("selMultiplicArea")) <> "" Then
	strMultiplic = trim(Request("selMultiplicArea"))
elseif trim(Request("selMultiplicComp")) <> "" Then
	strMultiplic = trim(Request("selMultiplicComp"))
end if

if Request("txtDtIni") <> "" Then
	strDtIni = Request("txtDtIni")
else
	strDtIni = ""
end if

if Request("txtDtFim") <> "" Then
	strDtFim = Request("txtDtFim")
else
	strDtFim = ""
end if

if Request("txtHrIni") <> "" Then
	strHrIni = Request("txtHrIni")
else
	strHrIni = "08:00"
end if

if Request("txtHrFim") <> "" Then
	strHrFim = Request("txtHrFim")
else
	strHrFim = "17:00"
end if 		

'Response.WRITE intCdCorte
'Response.END
strMandante		=  trim(Request("txtMandante"))

'*** Nome da Turma ***
strNomeTurma 	=  trim(Ucase(Request("txtNomeTurma")))

intCDsUnidades	= trim(Request("txtUnidades_Selecionadas"))

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

'*** REFERENTE A LISTA DAS UNIDADES
strSQLUnidade = ""
strSQLUnidade = strSQLUnidade & "SELECT UNID_CD_UNIDADE, "
strSQLUnidade = strSQLUnidade & "UNID_TX_DESC_UNIDADE "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE " 
strSQLUnidade = strSQLUnidade & "WHERE CORT_CD_CORTE = " & intCdCorte

if strAcao = "A" then	
	strTurma = trim(Request("pTurma"))	
	'*** PARA RETIRAR OS CADASTRADOS DA LISTA DE SELEÇÃO
	strSQLUnidade = strSQLUnidade & "AND UNID_CD_UNIDADE NOT IN " 
	strSQLUnidade = strSQLUnidade & "(SELECT UNIDADE.UNID_CD_UNIDADE "
	strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE UNIDADE, GRADE_TURMA_UNIDADE TUR_UNID "
	strSQLUnidade = strSQLUnidade & "WHERE UNIDADE.UNID_CD_UNIDADE = TUR_UNID.UNID_CD_UNIDADE "
	strSQLUnidade = strSQLUnidade & "AND UNIDADE.CORT_CD_CORTE = " & intCdCorte
	strSQLUnidade = strSQLUnidade & " AND TUR_UNID.CORT_CD_CORTE = " & intCdCorte
	strSQLUnidade = strSQLUnidade & " AND TUR_UNID.TURM_NR_CD_TURMA = " & strTurma & ")"
end if

strSQLUnidade = strSQLUnidade & " ORDER BY UNID_TX_DESC_UNIDADE"
'Response.write strSQLUnidade & "<br>"
'Response.end
set rsUnidade = db_banco.Execute(strSQLUnidade)

'Response.write "intCDsUnidades - " & intCDsUnidades

if intCDsUnidades <> "" then
	'*** REFERENTE A LISTA DAS UNIDADES - SELECIONADAS
	strSQLAltUnidade = ""
	strSQLAltUnidade = strSQLAltUnidade & "SELECT UNID_CD_UNIDADE, CORT_CD_CORTE, "
	strSQLAltUnidade = strSQLAltUnidade & "UNID_TX_DESC_UNIDADE "
	strSQLAltUnidade = strSQLAltUnidade & "FROM GRADE_UNIDADE "
	strSQLAltUnidade = strSQLAltUnidade & "WHERE CORT_CD_CORTE = " & intCdCorte 
	'Response.write strSQLAltUnidade
	'Response.end
	set rdsAltTumaUnidade = db_banco.Execute(strSQLAltUnidade)
	
	if not rdsAltTumaUnidade.EOF then			
		strMostraUnidade = "Sim"			
	end if		
end if		


'*** SELECIONA O NOME DA SALA ***
strSQLNomeSala = ""
strSQLNomeSala = strSQLNomeSala & "SELECT SALA_TX_NOME_SALA " 
strSQLNomeSala = strSQLNomeSala & "FROM GRADE_SALA "
strSQLNomeSala = strSQLNomeSala & "WHERE SALA_CD_SALA = " & strSala
strSQLNomeSala = strSQLNomeSala & " AND CORT_CD_CORTE = " & intCdCorte
set rdsNomeESala = db_banco.Execute(strSQLNomeSala)

if not rdsNomeESala.eof then
	strNomeSala = rdsNomeESala("SALA_TX_NOME_SALA")
else
	strNomeSala = ""
end if

rdsNomeESala.close
set rdsNomeESala = nothing
	
'*** MONTA O COMBO DE TURMA	
strSQLCurso = ""
strSQLCurso = strSQLCurso & "SELECT CURS_CD_CURSO, CURS_NUM_CARGA_CURSO, CURS_TX_METODO_CURSO "
strSQLCurso = strSQLCurso & "FROM GRADE_CURSO "
strSQLCurso = strSQLCurso & "WHERE CORT_CD_CORTE = " & intCdCorte
strSQLCurso = strSQLCurso & " ORDER BY CURS_CD_CURSO"	
'Response.write strSQLCurso

set rdsCurso = db_banco.Execute(strSQLCurso)

if strAcao = "A" then	

	strTurma = trim(Request("pTurma"))
		
	strSQLAltTurma = ""
	strSQLAltTurma = strSQLAltTurma & "SELECT TURM_NR_CD_TURMA, TURM_TX_DESC_TURMA, CURS_CD_CURSO, CORT_CD_CORTE,  "
	strSQLAltTurma = strSQLAltTurma & "USMA_CD_USUARIO, TURM_TX_MANDANTE, "
	strSQLAltTurma = strSQLAltTurma & "TURM_DT_INICIO, TURM_DT_TERMINO, TURM_HR_INICIO, "
	strSQLAltTurma = strSQLAltTurma & "TURM_HR_TERMINO, TURM_NUM_QTE_PERIODO "
	strSQLAltTurma = strSQLAltTurma & "FROM GRADE_TURMA "
	strSQLAltTurma = strSQLAltTurma & "WHERE SALA_CD_SALA =" & strSala
	strSQLAltTurma = strSQLAltTurma & " AND CORT_CD_CORTE = " & intCdCorte
	'Response.write  strSQLAltTurma
	
	Set rdsAltTurma = db_banco.Execute(strSQLAltTurma)			
	
	if not rdsAltTurma.EOF then
		strCurso 		 = rdsAltTurma("CURS_CD_CURSO")
		strMultiplic     = rdsAltTurma("USMA_CD_USUARIO")
		if trim(rdsAltTurma("TURM_DT_INICIO")) <> "" then
			strDtIni = MontaDataHora(trim(rdsAltTurma("TURM_DT_INICIO")),2)
		else
			strDtIni = ""
		end if
		
		if trim(rdsAltTurma("TURM_DT_TERMINO")) <> "" then
			strDtFim = MontaDataHora(trim(rdsAltTurma("TURM_DT_TERMINO")),2)
		else
			strDtFim = ""
		end if
		
		if trim(rdsAltTurma("TURM_HR_INICIO")) <> "" then
			strHrIni = MontaDataHora(trim(rdsAltTurma("TURM_HR_INICIO")),3)
		else
			strHrIni = "08:00"
		end if
		
		if trim(rdsAltTurma("TURM_HR_TERMINO")) <> "" then
			strHrFim = MontaDataHora(trim(rdsAltTurma("TURM_HR_TERMINO")),3)
		else
			strHrFim = "17:00"
		end if
		strMandante  = rdsAltTurma("TURM_TX_MANDANTE")
		'strCorte		 = rdsAltTurma("CORT_CD_CORTE")
		'strQtdePeriodo	 = rdsAltTurma("TURM_NUM_QTE_PERIODO")
		strNomeTurma 	 = trim(Ucase(rdsAltTurma("TURM_TX_DESC_TURMA")))
								
		'*** REFERENTE A LISTA DAS UNIDADES
		strSQLAltUnidade = ""
		strSQLAltUnidade = strSQLAltUnidade & "SELECT UNIDADE.UNID_TX_DESC_UNIDADE, "
		strSQLAltUnidade = strSQLAltUnidade & "UNIDADE.UNID_CD_UNIDADE "
		strSQLAltUnidade = strSQLAltUnidade & "FROM GRADE_UNIDADE UNIDADE, GRADE_TURMA_UNIDADE TUR_UNID "
		strSQLAltUnidade = strSQLAltUnidade & "WHERE UNIDADE.UNID_CD_UNIDADE = TUR_UNID.UNID_CD_UNIDADE "
		strSQLAltUnidade = strSQLAltUnidade & "AND UNIDADE.CORT_CD_CORTE = " & intCdCorte
		strSQLAltUnidade = strSQLAltUnidade & " AND TUR_UNID.CORT_CD_CORTE = " & intCdCorte
		strSQLAltUnidade = strSQLAltUnidade & " AND TUR_UNID.TURM_NR_CD_TURMA = " & strTurma
		'Response.write strSQLAltUnidade
		'Response.end
		set rdsAltTumaUnidade = db_banco.Execute(strSQLAltUnidade)
		
		if not rdsAltTumaUnidade.EOF then			
			strMostraUnidade = "Sim"			
		end if				
	end if
	
	rdsAltTurma.close
	set rdsAltTurma = nothing	

end if

'*** MONTA O COMBO DE MULTIPLICADOR	- UNIDADE
strSQLMult_Unid = ""
strSQLMult_Unid = strSQLMult_Unid & "SELECT MULT.MULT_NR_CD_CHAVE, MULT.MULT_NR_CD_ID_MULT, MULT.MULT_TX_NOME_MULTIPLICADOR "
strSQLMult_Unid = strSQLMult_Unid & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
strSQLMult_Unid = strSQLMult_Unid & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
strSQLMult_Unid = strSQLMult_Unid & "AND MULT.CORT_CD_CORTE = " & intCdCorte
strSQLMult_Unid = strSQLMult_Unid & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte

if strCurso <> "0" then
	strSQLMult_Unid = strSQLMult_Unid & " AND MULT_CURSO.CURS_CD_CURSO = '" & strCurso & "'"
else
	strSQLMult_Unid = strSQLMult_Unid & " AND MULT_CURSO.CURS_CD_CURSO = '999999'"
end if

strSQLMult_Unid = strSQLMult_Unid & " ORDER BY MULT.MULT_TX_NOME_MULTIPLICADOR"	
'Response.write strSQLMult_Unid & "<br>"
'Response.end
set rdsMultiplicadorUnid = db_banco.Execute(strSQLMult_Unid)


'*** MONTA O COMBO DE MULTIPLICADOR	- ÁREA DE NEGÓCIO
strSQLMult_AreNeg = ""
strSQLMult_AreNeg = strSQLMult_AreNeg & "SELECT MULT.MULT_NR_CD_CHAVE, MULT.MULT_NR_CD_ID_MULT, MULT.MULT_TX_NOME_MULTIPLICADOR "
strSQLMult_AreNeg = strSQLMult_AreNeg & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
strSQLMult_AreNeg = strSQLMult_AreNeg & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
strSQLMult_AreNeg = strSQLMult_AreNeg & "AND MULT.CORT_CD_CORTE = " & intCdCorte
strSQLMult_AreNeg = strSQLMult_AreNeg & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte

if strCurso <> "0" then
	strSQLMult_AreNeg = strSQLMult_AreNeg & " AND MULT_CURSO.CURS_CD_CURSO = '" & strCurso & "'"
else
	strSQLMult_AreNeg = strSQLMult_AreNeg & " AND MULT_CURSO.CURS_CD_CURSO = '999999'"
end if

strSQLMult_AreNeg = strSQLMult_AreNeg & " ORDER BY MULT.MULT_TX_NOME_MULTIPLICADOR"	
'Response.write strSQLMult_AreNeg & "<br>"
'Response.end
set rdsMultiplicadorArea = db_banco.Execute(strSQLMult_AreNeg)


'*** MONTA O COMBO DE MULTIPLICADOR	- COMPANHIA
strSQLMult_Companhia = ""
strSQLMult_Companhia = strSQLMult_Companhia & "SELECT MULT.MULT_NR_CD_CHAVE, MULT.MULT_NR_CD_ID_MULT, MULT.MULT_TX_NOME_MULTIPLICADOR "
strSQLMult_Companhia = strSQLMult_Companhia & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
strSQLMult_Companhia = strSQLMult_Companhia & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
strSQLMult_Companhia = strSQLMult_Companhia & "AND MULT.CORT_CD_CORTE = " & intCdCorte
strSQLMult_Companhia = strSQLMult_Companhia & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte

if strCurso <> "0" then
	strSQLMult_Companhia = strSQLMult_Companhia & " AND MULT_CURSO.CURS_CD_CURSO = '" & strCurso & "'"
else
	strSQLMult_Companhia = strSQLMult_Companhia & " AND MULT_CURSO.CURS_CD_CURSO = '999999'"
end if

strSQLMult_Companhia = strSQLMult_Companhia & " ORDER BY MULT.MULT_TX_NOME_MULTIPLICADOR"	
'Response.write strSQLMult_Companhia & "<br>"
'Response.end
set rdsMultiplicadorComp = db_banco.Execute(strSQLMult_Companhia)

on error resume next
	set temp = db_banco.execute("SELECT * FROM GRADE_TURMA WHERE SALA_CD_SALA=" & strSala & " AND TURM_NR_CD_TURMA=" & request("pTurma") & " AND CORT_CD_CORTE = " & intCdCorte)

	strMultiplic=temp("USMA_CD_USUARIO")
	strMandante=temp("TURM_TX_MANDANTE")
err.clear

'*************************************************   GRAVA REGISTRO DE TURMA  *******************************
if strGrava = "GravaTurma" then
	
		Erro = 0

		'====== VERIFICA O TOTAL DE PERÍODOS DO CURSO =========================================
		
		set rscurso = db_banco.execute("SELECT * FROM GRADE_CURSO WHERE CURS_CD_CURSO='" & strCurso & "' AND CORT_CD_CORTE = " & intCdCorte)
		rstCarga = rscurso("CURS_NUM_CARGA_CURSO")		
		rstPeriodos = rstCarga/4
		
		'====== VERIFICA SE OS PERÍODOS DO CURSO TEM COINCIDÊNCIA COM AS DATAS ESCOLHIDAS ===========
		
		Diferenca=DateDiff("d", strDtIni, strDtFim)
		
		Diferenca = Diferenca + 1
		
		strDias = rstPeriodos/2
		
		if Diferenca <> strDias then
			Erro = 1
		end if

		'====== VERIFICA SE AS DATAS ESCOLHIDAS SÃO MAIORES QUE A DATA ATUAL===========
		
		if cdate(strDtIni)  < cdate(date) or cdate(strDtFim) < cdate(date) then
			Erro = 6
		end if		
		
		'====== VERIFICA SE A FAIXA DE DATAS ESCOLHIDA É DE DIAS CORRIDOS ===========================

		if Erro=0 then
		
		verifica = 0
		corridos = 0
		
		data_inicial = strDtIni
		
		do until verifica = Diferenca 
		
			semana = WeekDay(data_inicial)
				
			if semana<>7 and semana<>1 then
				corridos = corridos + 1
			end if
		
			verifica = verifica + 1
			
			data_inicial = cdate(data_inicial) + 1

		loop

		if corridos <> Diferenca then
			Erro = 2
		end if

		End if		
		
		'====== VERIFICA SE EXISTE ALGUM FERIADO NACIONAL / REGIONAL NA FAIXA DE DATAS ESCOLHIDA ==========
		
		if Erro=0 then

		verifica = 0
		nao_feriado = 0
		
		data_inicial = strDtIni
		
		do until verifica = Diferenca 
		
			ver_feriado = (right("000" & day(data_inicial),2) & "/" & right("000" & month(data_inicial),2))
			
			ssql="SELECT * FROM GRADE_FERIADO WHERE FERI_TX_TIPO_FERIADO='0' AND FERI_DT_DATA_FERIADO='" & ver_feriado & "'"
			
			set temp = db_banco.execute(ssql)
			
			if temp.eof=true then
			
				ssql=""
				ssql="SELECT GRADE_FERIADO_SALA.SALA_CD_SALA, "
				ssql = ssql + "GRADE_FERIADO_SALA.FERI_CD_FERIADO, GRADE_FERIADO.FERI_TX_NOME_FERIADO, GRADE_FERIADO.FERI_DT_DATA_FERIADO "
				ssql = ssql + "FROM  GRADE_FERIADO_SALA INNER JOIN GRADE_FERIADO ON "
				ssql = ssql + "GRADE_FERIADO_SALA.FERI_CD_FERIADO = GRADE_FERIADO.FERI_CD_FERIADO "
				ssql = ssql + "WHERE GRADE_FERIADO_SALA.SALA_CD_SALA = " & strSala & " "
				ssql = ssql + "AND GRADE_FERIADO.FERI_DT_DATA_FERIADO = '" & ver_feriado & "' "
				ssql = ssql + "AND GRADE_FERIADO_SALA.CORT_CD_CORTE = " & intCdCorte
				
				set temp = db_banco.execute(ssql)
			
			end if
				
			if temp.eof=true then
				nao_feriado = nao_feriado + 1
			end if
		
			verifica = verifica + 1
			
			data_inicial = cdate(data_inicial) + 1

		loop
				
		if trim(nao_feriado) <> trim(Diferenca) then
			Erro = 3
		end if
		
		End if		

		'====== VERIFICA SE A SALA ESTÁ OCUPADA NO PERÍODO ===================================================
		
		if Erro=0 then		
		
			inicio = year(strDtIni) & "-" & right("000" & month(strDtIni),2) & "-" & right("000" & day(strDtIni),2)
			Data_fim = cdate(strDtFim)+1
			fim = year(Data_fim) & "-" & right("000" & month(Data_fim),2) & "-" & right("000" & day(Data_fim),2)
		
			ssql=""
			ssql="SELECT     * "
			ssql=ssql + "FROM         dbo.GRADE_TURMA "
			ssql=ssql + "WHERE     (SALA_CD_SALA = " & strSala & ") AND "
			ssql=ssql + "(TURM_DT_INICIO BETWEEN CONVERT(DATETIME, '" & inicio & " 00:00:00', 102) AND CONVERT(DATETIME,'" & fim & " 00:00:00', 102)) "
			ssql=ssql + "AND CORT_CD_CORTE = " & intCdCorte
		
			set rsinicio = db_banco.execute(ssql)

			ssql=""
			ssql="SELECT     * "
			ssql=ssql + "FROM         dbo.GRADE_TURMA "
			ssql=ssql + "WHERE     (SALA_CD_SALA = " & strSala & ") AND "
			ssql=ssql + "(TURM_DT_TERMINO BETWEEN CONVERT(DATETIME, '" & inicio & " 00:00:00', 102) AND CONVERT(DATETIME,'" & fim & " 00:00:00', 102)) "
			ssql=ssql + "AND CORT_CD_CORTE = " & intCdCorte
		
			set rstermino = db_banco.execute(ssql)
		
			if rsinicio.eof=false or rstermino.eof=false then
				Erro = 4
			end if
			
			if Erro=0 then
			
				tem_turma = 0
				
				set turmas = db_banco.execute("SELECT * FROM GRADE_TURMA WHERE SALA_CD_SALA=" & strSala & " AND CORT_CD_CORTE = " & intCdCorte)
				
				
				inicial = cdate(strDtIni)
				final = cdate(strDtFim)
								
				do until turmas.eof=true
				
					data_inicio_turma = cdate(turmas("TURM_DT_INICIO"))
					data_fim_turma = cdate(turmas("TURM_DT_TERMINO"))
					
					diferenca_datas = DateDiff("d",data_inicio_turma,data_fim_turma)
					
					valida_datas = 0
					
					do until valida_datas = diferenca_datas
					
						'response.write data_inicio_turma & "<p>"
						'response.write data_fim_turma & "<p>"					
						
						if data_inicio_turma = inicial or data_fim_turma = inicial or data_inicio_turma = final or data_fim_turma = final then
							tem_turma = tem_turma + 1						
						end if

						valida_datas = valida_datas + 1
						data_inicio_turma = cdate(data_inicio_turma) + 1
						data_fim_turma = cdate(data_fim_turma) - 1

					loop
				
					turmas.movenext
				loop
				
				'response.write tem_turma & "<p>"

			if tem_turma<>0 then
				Erro = 4
			end if
							
			end if

		end if

		'====== VERIFICA SE O MULTIPLICADOR ESTÁ OCUPADO NO PERÍODO ==========================================
		
		if Erro=0 then		
		
			inicio = year(strDtIni) & "-" & right("000" & month(strDtIni),2) & "-" & right("000" & day(strDtIni),2)
			Data_fim = cdate(strDtFim)+1
			fim = year(Data_fim) & "-" & right("000" & month(Data_fim),2) & "-" & right("000" & day(Data_fim),2)
		
			ssql=""
			ssql="SELECT     * "
			ssql=ssql + "FROM         dbo.GRADE_TURMA "
			ssql=ssql + "WHERE     (USMA_CD_USUARIO = '" & strMultiplic & "') AND "
			ssql=ssql + "(TURM_DT_INICIO BETWEEN CONVERT(DATETIME, '" & inicio & " 00:00:00', 102) AND CONVERT(DATETIME,'" & fim & " 00:00:00', 102)) "
			ssql=ssql + "AND CORT_CD_CORTE = " & intCdCorte
			
			set rsinicio = db_banco.execute(ssql)

			ssql=""
			ssql="SELECT     * "
			ssql=ssql + "FROM         dbo.GRADE_TURMA "
			ssql=ssql + "WHERE     (USMA_CD_USUARIO = '" & strMultiplic & "') AND "
			ssql=ssql + "(TURM_DT_TERMINO BETWEEN CONVERT(DATETIME, '" & inicio & " 00:00:00', 102) AND CONVERT(DATETIME,'" & fim & " 00:00:00', 102)) "
			ssql=ssql + "AND CORT_CD_CORTE = " & intCdCorte
		
			set rstermino = db_banco.execute(ssql)
		
			if rsinicio.eof=false or rstermino.eof=false then
				Erro = 5
			end if

			if Erro=0 then
			
				tem_turma = 0
				
				set turmas = db_banco.execute("SELECT * FROM GRADE_TURMA WHERE SALA_CD_SALA=" & strSala & " AND CORT_CD_CORTE = " & intCdCorte)
				
				inicial = cdate(strDtIni)
				final = cdate(strDtFim)
								
				do until turmas.eof=true
				
					data_inicio_turma = cdate(turmas("TURM_DT_INICIO"))
					data_fim_turma = cdate(turmas("TURM_DT_TERMINO"))
					multiplicador = turmas("USMA_CD_USUARIO")
					
					diferenca_datas = DateDiff("d",data_inicio_turma,data_fim_turma)
					
					valida_datas = 0
					
					do until valida_datas = diferenca_datas
					
						'response.write data_inicio_turma & "<p>"
						'response.write data_fim_turma & "<p>"					
						
						if (data_inicio_turma = inicial or data_fim_turma = inicial or data_inicio_turma = final or data_fim_turma = final) AND (strMultiplic = multiplicador) then
							tem_turma = tem_turma + 1						
						end if

						valida_datas = valida_datas + 1
						data_inicio_turma = cdate(data_inicio_turma) + 1
						data_fim_turma = cdate(data_fim_turma) - 1

					loop
				
					turmas.movenext
					
				loop
				
				'response.write tem_turma & "<p>"

			if tem_turma<>0 then
				Erro = 5
			end if
							
			end if

		end if
		
		'====================== VERIFICA SE A DATA DE INICIO É MENOR QUE A DATA DE TÉRMINO DO MATERIAL DIDÁTICO ======
		
		ssql=""
		ssql="SELECT * FROM GRADE_CURSO WHERE CURS_CD_CURSO='" & strCurso & "' AND CORT_CD_CORTE = " & intCdCorte
		
		set datacurso = db_banco.execute(ssql)
		
		diferenca_datas = DateDiff("d", strDtIni, datacurso("CURS_DT_FIM_MATERIAL_DIDATICO"))
	
		if diferenca_datas < 10 then
				Erro = 7 
		end if
				
		'====================== VERIFICA SE EXITEM ERROS E EXIBE MENSAGEM =====================================
			
		select case Erro
		Case 1
			strMSG = "A faixa de Datas Escolhida não coincide com a carga horária do Curso, que é de " & rstCarga & " Hs"		
			incluir_turma = 0
		Case 2
			strMSG = "A faixa de Datas Escolhida não é de dias corridos ou é Fim de Semana!"		
			incluir_turma = 0
		Case 3
			strMSG = "Existem Feriados durante a Faixa de Datas Escolhida"		
			incluir_turma = 0
		Case 4
			strMSG = "A Sala encontra-se ocupada no decorrer da Faixa de Datas Escolhida"		
			incluir_turma = 0	
		Case 5
			strMSG = "O Multiplicador encontra-se ocupado no decorrer da Faixa de Datas Escolhida"		
			incluir_turma = 0
		Case 6
			strMSG = "As Datas selecionadas devem ser maiores que a data atual"		
			incluir_turma = 0
		Case 7
			strMSG = "Material Didático Não Concluído"		
			incluir_turma = 0		
		Case 0
			strMSG = ""
			incluir_turma = 1
		End select
		
		if incluir_turma = 1 then
		
			ins_inicio = year(strDtIni) & "-" & right("000" & month(strDtIni),2) & "-" & right("000" & day(strDtIni),2)
			ins_fim = year(Data_fim) & "-" & right("000" & month(Data_fim),2) & "-" & right("000" & day(Data_fim),2)
				
			set temp = db_banco.execute("SELECT MAX(TURM_NR_CD_TURMA) AS CODIGO FROM GRADE_TURMA WHERE CORT_CD_CORTE = " & intCdCorte)
							
			if not temp.eof then
				if isnull(temp("CODIGO")) then
					intCdTurma = 1
				else
					intCdTurma = temp("CODIGO") + 1
				end if
			else
				intCdTurma = 1
			end if	
								
			'set rsChaveMult = db_banco.execute("SELECT MULT_NR_CD_CHAVE FROM GRADE_MULTIPLICADOR WHERE CORT_CD_CORTE = " & intCdCorte & " AND MULT_NR_CD_ID_MULT = " & strMultiplic)	
							
			'if not rsChaveMult.eof then
				'strChaveMult = rsChaveMult("MULT_NR_CD_CHAVE")
			'else
				'strChaveMult = ""
			'end if				
																												
			strSQLIncTurma = ""
			strSQLIncTurma = strSQLIncTurma & "INSERT INTO GRADE_TURMA (CORT_CD_CORTE, TURM_NR_CD_TURMA, SALA_CD_SALA, CURS_CD_CURSO, MULT_NR_CD_ID_MULT, "	'*** Mudou - USMA_CD_USUARIO
			strSQLIncTurma = strSQLIncTurma & "TURM_TX_DESC_TURMA, TURM_TX_MANDANTE, TURM_DT_INICIO, TURM_DT_TERMINO, TURM_HR_INICIO, "
			strSQLIncTurma = strSQLIncTurma & "TURM_HR_TERMINO, TURM_NUM_QTE_PERIODO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
			strSQLIncTurma = strSQLIncTurma & "VALUES(" & intCdCorte & "," & intCdTurma & "," & strSala & ",'" & strCurso & "','" & strMultiplic & "','" 
			strSQLIncTurma = strSQLIncTurma & strNomeTurma & "','" & strMandante & "',"
			strSQLIncTurma = strSQLIncTurma & "CONVERT(DATETIME, '" & ins_inicio & " 00:00:00', 102), CONVERT(DATETIME, '" & ins_fim & " 00:00:00', 102)," 
			strSQLIncTurma = strSQLIncTurma & "CONVERT(DATETIME, '1899-12-30 08:00:00', 102), CONVERT(DATETIME, '1899-12-30 17:00:00', 102)," 
			strSQLIncTurma = strSQLIncTurma & rstPeriodos & ",'I','" & Session("CdUsuario") & "',GETDATE())"	
			'response.write strSQLIncTurma
		    'Response.end
			
		  	'*** LIMPA OS REGISTROS DA TABELA GRADE_TURMA_UNIDADE **
			'strSQLDelTurmaUnid = ""
			'strSQLDelTurmaUnid = strSQLDelTurmaUnid & "DELETE FROM GRADE_TURMA_UNIDADE "	
			'strSQLDelTurmaUnid = strSQLDelTurmaUnid & "WHERE FERI_CD_FERIADO = " & intCdFeriado 
			'strSQLDelTurmaUnid = strSQLDelTurmaUnid & " AND CORT_CD_CORTE = " & Session("Corte")
			'response.write strSQLDelTurmaUnid
			'Response.end		
			
			'db_banco.Execute(strSQLDelTurmaUnid)	
			
			'*** CADASTRA AS NOVOS REGISTROS EM GRADE_TURMA_UNIDADE ***			
			vetCDsUnidades = split(intCDsUnidades,",")
						
			r = 0			
			for r = lbound(vetCDsUnidades) to Ubound(vetCDsUnidades)				
				
				'Response.write vetCDsUnidades(r) & "<br><br>"		
				
				if vetCDsUnidades(r) <> "" then				
					intCdUnidadeResult = cint(vetCDsUnidades(r))
							
					strSQLIncTurmaUnid = ""		
					strSQLIncTurmaUnid = strSQLIncTurmaUnid & "INSERT INTO GRADE_TURMA_UNIDADE "
					strSQLIncTurmaUnid = strSQLIncTurmaUnid & "(CORT_CD_CORTE, TURM_NR_CD_TURMA, UNID_CD_UNIDADE, "
					strSQLIncTurmaUnid = strSQLIncTurmaUnid & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
					strSQLIncTurmaUnid = strSQLIncTurmaUnid & "VALUES(" & intCdCorte & "," & intCdTurma & "," & intCdUnidadeResult & "," 
					strSQLIncTurmaUnid = strSQLIncTurmaUnid & "'I','" & Session("CdUsuario") & "',GETDATE())"					
					'response.write strSQLIncTurmaUnid & "<br><br>"
					'Response.end					
					db_banco.Execute(strSQLIncTurmaUnid)
				end if
			next			

			on error resume next
				db_banco.Execute strSQLIncTurma
	
			if err.number = 0 then		
				strMSG = "A turma - " & Ucase(strNomeTurma) & " foi incluída com sucesso."
				
				strNomeTurma	= ""
				strUnidDir		= ""
				strCurso 		= 0
				strMultiplic	= 0
				strMandante		= ""
				strDtIni 		= ""
				strDtFim 		= ""			
				strHrIni 		= "08:00"
				strHrFim 		= "17:00"			
			else
				strMSG = "Houve um erro na inclusão da turma - " & Ucase(strNomeTurma) & " - " & err.description
			end if

		end if
		
end if
	

'****************************** FIM DA ROTINA DE INCLUSÃO ******************************************

public function MontaDataHora(strData,intDataTime)

	'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
	'*** intDataTime = 1 (DATA E HORA)
	'*** intDataTime = 2 (DATA)
	'*** intDataTime = 3 (HORA)
	'*** intDataTime = 4 (FORMATO DE BANCO)
	'*** intDataTime = 5 (FORMATO DE BANCO - DIA E MÊS)

	if day(strData) < 10 then
		strDia = "0" & day(strData)		
	else
		strDia = day(strData)		
	end if
	
	if month(strData) < 10 then
		strMes = "0" & month(strData)	
	else
		strMes = month(strData)	
	end if		
	
	if hour(strData) < 10 then
		strHora = "0" & hour(strData)		
	else
		strHora = hour(strData)		
	end if
	
	if minute(strData) < 10 then
		strMinuto = "0" & minute(strData)	
	else
		strMinuto = minute(strData)	
	end if	

	if cint(intDataTime) = 1 then	
		MontaDataHora = strDia & "/" & strMes & "/" & year(strData) & " - " &  strHora & ":" & strMinuto	
	elseif cint(intDataTime) = 2 then	
		MontaDataHora = strDia & "/" & strMes & "/" & year(strData) 
	elseif cint(intDataTime) = 3 then	
		MontaDataHora = strHora & ":" & strMinuto	
	elseif cint(intDataTime) = 4 then
		MontaDataHora = strMes & "/" & strDia & "/" & year(strData)
	elseif cint(intDataTime) = 5 then
		MontaDataHora = strDia & "/" & strMes
	end if
end function

%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript" src="../js/troca_lista.js"></script>
		
		<script language="javascript">
			function Confirma()
			{							
				if(document.frmCadTurma.txtNomeTurma.value == "")
				{
				alert("É necessário o preenchimento do campo TURMA!");
				document.frmCadTurma.txtNomeTurma.focus();
				return;
				}
								
				if(document.frmCadTurma.selCurso.selectedIndex == 0)
				{
				alert("Selecione um CURSO!");
				document.frmCadTurma.selCurso.focus();
				return;
				}
												
				if(document.frmCadTurma.selMultiplic.selectedIndex == 0)
				{
				alert("Selecione um MULTIPLICADOR!");
				document.frmCadTurma.selMultiplic.focus();
				return;
				}
				
				if(document.frmCadTurma.txtDtIni.value == "")
				{
				alert("É necessário o preenchimento do campo DATA DE INÍCIO!");
				document.frmCadTurma.txtDtIni.focus();
				return;
				}			
				
				if(document.frmCadTurma.txtDtFim.value == "")
				{
				alert("É necessário o preenchimento do campo DATA DE TÉRMINO!");
				document.frmCadTurma.txtDtFim.focus();
				return;
				}						
								
				if (document.frmCadTurma.parAcao.value == 'I')
				{ 
					document.frmCadTurma.action="cadastra_turma.asp?parGrava=GravaTurma"
				}
				
				if (document.frmCadTurma.parAcao.value == 'A')
				
				{ 				
					document.frmCadTurma.action="grava_turma.asp"
				}
				
				//*** Monta uma string com as Unidades Selecionadss, separadas por vírgula
				carrega_txt(document.frmCadTurma.selUnidade_Selecionado)
				
				document.frmCadTurma.submit();			
			}
			
			function carrega_txt(fbox) 
			{				
				document.frmCadTurma.txtUnidades_Selecionadas.value = '';
				for(var i=0; i<fbox.options.length; i++) 
				{
					if (i == 0)
					{
						document.frmCadTurma.txtUnidades_Selecionadas.value = fbox.options[i].value;
					}
					else
					{					
						document.frmCadTurma.txtUnidades_Selecionadas.value = document.frmCadTurma.txtUnidades_Selecionadas.value + "," + fbox.options[i].value;
					}
				}
			}
			
			function submet_pagina(strValor, strTipo)
			{						
				var strAcao = document.frmCadTurma.parAcao.value;				
				var strCdSala = document.frmCadTurma.hdSala.value;				
				var strNomeTurma = document.frmCadTurma.txtNomeTurma.value;				
				var strMandante = document.frmCadTurma.txtMandante.value;				
				var strDtIni = document.frmCadTurma.txtDtIni.value;				
				var strDtFim = document.frmCadTurma.txtDtFim.value;				
				var strHrIni = document.frmCadTurma.txtHrIni.value;				
				var strHrFim = document.frmCadTurma.txtHrFim.value;	
				var strCorte = document.frmCadTurma.selCorte.value;	
				var strCurso = document.frmCadTurma.selCurso.value;	
				var strDiasCurso = document.frmCadTurma.hdDiasCurso.value;	
				
				if (strTipo == 'Curso')		
				{				
					alert("Entrou no Curso -" + document.frmCadTurma.txtUnidades_Selecionadas.value);
				
					document.frmCadTurma.txtUnidades_Selecionadas.value = '';
					
					//*** Monta uma string com as Unidades Selecionadss, separadas por vírgula
					carrega_txt(document.frmCadTurma.selUnidade_Selecionado)
					
					alert(strUnidSelecionadas);
					
					var strUnidSelecionadas = document.frmCadTurma.txtUnidades_Selecionadas.value	
																	
					window.location.href='cadastra_turma.asp?parAcao='+strAcao+'&hdSala='+strCdSala+'&selCurso='+strValor+'&txtNomeTurma='+strNomeTurma+'&txtMandante='+strMandante+'&txtDtIni='+strDtIni+'&txtDtFim='+strDtFim+'&txtHrIni='+strHrIni+'&txtHrFim='+strHrFim+'&pCorte='+strCorte+'&txtUnidades_Selecionadas='+strUnidSelecionadas+'&hdDiasCurso='+strDiasCurso;										
				}
				
				if (strTipo == 'Unidade')		
				{
					alert("Entrou no Unidade -" + document.frmCadTurma.txtUnidades_Selecionadas.value);
					
					document.frmCadTurma.txtUnidades_Selecionadas.value = '';
					
					//*** Monta uma string com as Unidades Selecionadss, separadas por vírgula
					carrega_txt(document.frmCadTurma.selUnidade_Selecionado)
					
					var strUnidSelecionadas = document.frmCadTurma.txtUnidades_Selecionadas.value	
					
					alert(strUnidSelecionadas);
					
					window.location.href='cadastra_turma.asp?parAcao='+strAcao+'&hdSala='+strCdSala+'&selCurso='+strCurso+'&txtNomeTurma='+strNomeTurma+'&txtMandante='+strMandante+'&txtDtIni='+strDtIni+'&txtDtFim='+strDtFim+'&txtHrIni='+strHrIni+'&txtHrFim='+strHrFim+'&pCorte='+strCorte+'&txtUnidades_Selecionadas='+strUnidSelecionadas;
				}
			}
			
			function MostraEscondeMult()
			{											
				if (document.frmCadTurma.rdMultplicador(0).checked == true)
				{
					document.frmCadTurma.selMultiplic.disabled = false;					
					document.frmCadTurma.selMultiplicArea.selectedIndex = 0;		
					document.frmCadTurma.selMultiplicArea.disabled = true;
					document.frmCadTurma.selMultiplicComp.selectedIndex = 0;					
					document.frmCadTurma.selMultiplicComp.disabled = true;			
				}
				
				if (document.frmCadTurma.rdMultplicador(1).checked == true)
				{
					document.frmCadTurma.selMultiplic.selectedIndex = 0;
					document.frmCadTurma.selMultiplic.disabled = true;			
					document.frmCadTurma.selMultiplicArea.disabled = false;
					document.frmCadTurma.selMultiplicComp.selectedIndex = 0;
					document.frmCadTurma.selMultiplicComp.disabled = true;			
				}
				
				if (document.frmCadTurma.rdMultplicador(2).checked == true)
				{
					document.frmCadTurma.selMultiplic.selectedIndex = 0;
					document.frmCadTurma.selMultiplic.disabled = true;		
					document.frmCadTurma.selMultiplicArea.selectedIndex = 0;	
					document.frmCadTurma.selMultiplicArea.disabled = true;
					document.frmCadTurma.selMultiplicComp.disabled = false;			
				}
			}
			
			function VerificaUnidades(strTipo)
			{								
				if (document.frmCadTurma.selUnidade_Selecionado.options.length == 0)
				{
					if (strTipo == 'Unidade')
					{
						alert("Para a opção MULTIPLICADOR - UNIDADE ,é necessária a seleção de plo menos uma UNIDADE!");
						document.frmCadTurma.selUnidade.focus();
						
						document.frmCadTurma.selMultiplic.selectedIndex = 0;
						document.frmCadTurma.selMultiplic.disabled = true;		
						document.frmCadTurma.selMultiplicArea.selectedIndex = 0;	
						document.frmCadTurma.selMultiplicArea.disabled = true;
						document.frmCadTurma.selMultiplicComp.disabled = false;			
						return;
					}
					
					if (strTipo == 'Area')
					{
						alert("Para a opção MULTIPLICADOR - ÁREA DE NEGÓCIO ,é necessária a seleção de plo menos uma UNIDADE!");
						document.frmCadTurma.selUnidade.focus();
						
						document.frmCadTurma.selMultiplic.selectedIndex = 0;
						document.frmCadTurma.selMultiplic.disabled = true;		
						document.frmCadTurma.selMultiplicArea.selectedIndex = 0;	
						document.frmCadTurma.selMultiplicArea.disabled = true;
						document.frmCadTurma.selMultiplicComp.disabled = false;							
						return;
					}
				}			
			}
			
			function VerificaCurso(strTipo)
			{
				if (document.frmCadTurma.selCurso.selectedIndex == 0)
				{
					alert("Selecione um CURSO para listar os MULTIPLICADORES!");
					document.frmCadTurma.selCurso.focus();
					return;
				}			
			}
			
			function CarregaDataFim(strDtInicio)
			{
				alert(strDtInicio + ' - ' + document.frmCadTurma.hdDiasCurso.value);			
			}
			
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadTurma">	
		  
			<input type="hidden" value="<%=strSala%>" name="hdSala"> 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 
			<input type="hidden" value="Turma" name="parTipo"> 
			<input type="hidden" value="<%=Session("Corte")%>" name="selCorte"> 
			<input type="hidden" name="txtUnidades_Selecionadas">
			<input type="text" name="hdDiasCurso" value="<%=strDias%>">								
									
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
						<div align="center"><a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr bgcolor="#F1F1F1">
				<td colspan="3" height="20">
				  
        <table width="859" border="0" align="center">
          <tr> 
            <td width="92"> 
              <div align="right"><a href="javascript:Confirma();"><img border="0" src="../../imagens/confirma_f02.gif"></a></div>
            </td>
            <td width="46"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
            <td width="28">&nbsp;</td>
            <td width="81"></td>
            <td width="24"><a href="inclui_altera_turma.asp?selSala=<%=strSala%>"><img border="0" src="../../imagens/volta_f02.gif" title="Volta para a Tela de Cadastro de Sala e Turma - Grade de Treinamento"></a></td>
            <td width="211"><font color="#330099" face="Verdana" size="2"><b>Tela de Sala e Turma</b></font></td>			 
            <td width="26"></td>
		    <td width="317"></td>
		  </tr>
        </table>
				</td>
			  </tr>
			</table>
					
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td height="10">
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Turmas - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td height="10"></td>
				</tr>
			  </table>
			  <table border="0" width="988" height="534">
			  
			  	<tr>
			  	  <td height="20"></td>
			  	  <td height="20" valign="middle" align="center" colspan="2">
				  	<%if strGrava = "GravaTurma" then%>
				  	<font face="Verdana" color="#FE5A31" size="2"><b><%=strMSG%></b></font>
				  <%end if%>				  
				  </td>	
				  <td height="20"></td>				 
		  	    </tr>
				
			  	<tr>
			  	  <td height="21"></td>			  	 
			  	  <td height="21" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=strNomeAcao%></font></td>
				  <td height="20"></td>						 
				</tr>
			  
				<%
				strSQLCorte = ""
				strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
				strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
				strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
				'Response.write strSQLCorte
				'Response.end
				set rsCorte = db_banco.Execute(strSQLCorte)
			 
				if not rsCorte.eof then
					strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & 	rsCorte("CORT_DT_DATA_CORTE")					
				else
					strNomeCorte = ""
				end if
				
				rsCorte.close
				set rsCorte = nothing			 
				%>				   
			  	<tr>
			  	  <td height="7" colspan="3"></td>	
				   <td width="237" height="27" rowspan="11" align="left" valign="top">				  	
					<table width="100%" border="1" bordercolor="#FF3300" cellspacing="0" cellpadding="0">
					  <tr>
						<td>
							<table width="100%"  border="0" cellspacing="0" cellpadding="0">
							  <tr>
							  	<td>&nbsp;</td>
								<td>
									<p align="center">
										<font face="Verdana" color="#FF3300" size="1"><b>Verificações para Cadastro:</b></font>
									</p>
									<p align="justify">
										<font face="Verdana" color="#FF3300" size="1">
											&nbsp;&nbsp;&nbsp;- Data de término é preenchida automaticamente de acordo com a data de início escolhida e carga horária do curso.<br><br>
											&nbsp;&nbsp;&nbsp;- Período do curso coinscidente com feriados ou finais de senama.<br><br>		
											&nbsp;&nbsp;&nbsp;- Sala escolhida já ocupada no período.<br><br>		
											&nbsp;&nbsp;&nbsp;- Multiplicador já alocado para o mesmo período.<br><br>			
											&nbsp;&nbsp;&nbsp;- Lista de multiplicadores contendo somente os já treinados.<br><br>			
											&nbsp;&nbsp;&nbsp;- Só permite confirmação de turma para materias didáticos já prontos na data de início de turma.<br><br>			
											&nbsp;&nbsp;&nbsp;- A data de início da turma não pode ser maior do que a data atual.<br>
											<br>							
									  </font>	
									</p>										
								</td>								
								<td>&nbsp;</td>
							  </tr>
							</table>
						</td>
					  </tr>
					</table>
				  </td>		  	  			 	  	 
		  	    </tr>
			  	<tr>
			  	  <td height="7"></td>
			  	  <td height="7" valign="middle" align="left"> <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font></td>
			  	  <td height="7" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><%=strNomeCorte%></font></td>
	  	        </tr>
			  					
			  	<tr>
			  	  <td height="27"></td>
			  	  <td height="27" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Sala:</b></font></td>
			  	  <td height="27" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><%=strNomeSala%></font></td>		  	 
				</tr>
				
				<tr> 
				  <td width="1" height="31"></td>
				  <td width="136" height="31" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Turma:</b></font></td>
				  <td height="31" valign="middle" align="left" width="596">
				  	<%'if strNomeTurma <> "" and strAcao = "A" then%>						
						<input type="hidden" name="txtCDTurma" value="<%=strTurma%>">	
					<%'else%>
						<input type="text" name="txtNomeTurma" maxlength="50" size="50" value="<%=strNomeTurma%>">	
					<%'end if%>						  
				  </td>
				</tr> 
				
			  	<tr> 
				  <td width="1" height="31"></td>
				  <td width="136" height="31" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso:</b></font></td>
				  <td height="31" valign="middle" align="left" width="596"> 
					<!--<input type="text" name="txtCurso" maxlength="100" size="50" value="<%'=strCurso%>">-->
					<select size="1" name="selCurso" onchange="javascript:submet_pagina(this.value,'Curso');">
					  <option value="0">== Selecione o Curso ==</option>
						<%
						do until rdsCurso.eof = true
							  if trim(strCurso) = trim(rdsCurso("CURS_CD_CURSO")) then%>
								  <option value="<%=rdsCurso("CURS_CD_CURSO")%>" selected><%=rdsCurso("CURS_CD_CURSO")%> - <%=rdsCurso("CURS_NUM_CARGA_CURSO")%> Hs - <%=rdsCurso("CURS_TX_METODO_CURSO")%></option>
							  <%else%>
									<option value="<%=rdsCurso("CURS_CD_CURSO")%>"><%=rdsCurso("CURS_CD_CURSO")%> - <%=rdsCurso("CURS_NUM_CARGA_CURSO")%> Hs - <%=rdsCurso("CURS_TX_METODO_CURSO")%></option>
							  <%end if						
							rdsCurso.movenext
						loop
						
						rdsCurso.close
						set rdsCurso = nothing						
						%>
					</select>				  
				  </td>
				</tr> 														
				 <tr>
				   <td height="165"></td>
				   <td height="165" valign="top" align="left" colspan="2">
				   
				   	 <table width="732" height="154" border="0">
						<tr>
							<td height="18" colspan="3"><font face="Verdana" size="2" color="#330099"><b>Unidade:</b></font></td>
						</tr>
						<tr> 
						  <td width="132" height="55" rowspan="5" align="center" valign="middle">        
						  <p align="left">
							<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dispon&iacute;veis:</font></p>
							<p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
							
							  <select name="selUnidade" size="5" multiple>									  
								<%
								if intCDsUnidades <> "" then
									
									'vetCDsUnid = split(intCDsUnidades,",")
									
									w = 0	
									do while not rsUnidade.eof										
										
										'for w = lbound(vetCDsUnid) to ubound(vetCDsUnid)
											'if trim(vetCDsUnid(w)) <> "" then
											
											'str_vet = trim(vetCDsUnid(w))
											'str_unid = trim(rsUnidade("UNID_CD_UNIDADE"))
											'if str_vet <> str_unid then
												
											if InStr(1, intCDsUnidades, rsUnidade("UNID_CD_UNIDADE")) = 0 then
											
												'if trim(vetCDsUnid(w)) <> trim(rsUnidade("UNID_CD_UNIDADE")) then
												%>
													<option value="<%=rsUnidade("UNID_CD_UNIDADE") & " - " & rsUnidade("CORT_CD_CORTE")%>"><%'=cint(vetCDsUnid(j)) & " <> " & cint(rsUnidade("UNID_CD_UNIDADE")) & " - "%><%=rsUnidade("UNID_TX_DESC_UNIDADE")%></option>
												<%		
												'end if		
											end if				
										'next	
										rsUnidade.movenext
									loop		
								else
									do until rsUnidade.eof = true
										%>
										<option value="<%=rsUnidade("UNID_CD_UNIDADE")%>"><%=rsUnidade("UNID_TX_DESC_UNIDADE")%></option>
										<%
										rsUnidade.movenext
									loop
								end if
								%>
							</select>
						</font>
							</p>
						  </td>
						  <td width="103" height="32" align="center" valign="middle"><div align="left"></div></td>
						  <td width="516" rowspan="5" align="center" valign="middle">
						  <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Selecionadas:</font></p>
							<p align="left">					
								<select name="selUnidade_Selecionado" size="5" multiple>
								<%
								if strMostraUnidade = "Sim" then										
									
									if intCDsUnidades <> "" then
										
										vetCDsUnidadesSel = split(intCDsUnidades,",")
										
										i = 0
										do while not rdsAltTumaUnidade.eof					
																						
											for i = lbound(vetCDsUnidadesSel) to ubound(vetCDsUnidadesSel)
												if cint(vetCDsUnidadesSel(i)) <> "" then
													if cint(vetCDsUnidadesSel(i)) = cint(rdsAltTumaUnidade("UNID_CD_UNIDADE")) then
													%>
														<option value="<%=rdsAltTumaUnidade("UNID_CD_UNIDADE")%>"><%=rdsAltTumaUnidade("UNID_TX_DESC_UNIDADE")%></option>			
													<%		
													end if	
												end if				
											next
											
											rdsAltTumaUnidade.movenext
										loop											
									else									
										do while not rdsAltTumaUnidade.eof					
											%>
											<option value="<%=rdsAltTumaUnidade("UNID_CD_UNIDADE")%>"><%=rdsAltTumaUnidade("UNID_TX_DESC_UNIDADE")%></option>			
											<%								
											rdsAltTumaUnidade.movenext
										loop		
									end if							
						
									rdsAltTumaUnidade.close
									set rdsAltTumaUnidade = nothing	
									
										
								end if		
								%>					 
								</select>
							</p></td>
						</tr>
						
						<tr>
						  <td height="39" align="center" valign="middle"><div align="center"><img src="../../imagens/continua_F01.gif" width="24" height="24" onClick="move(document.frmCadTurma.selUnidade,document.frmCadTurma.selUnidade_Selecionado,1);submet_pagina('','Unidade');"></div></td>
						</tr>
						<tr>
						  <td height="26" align="center" valign="middle"><div align="center"><img src="../../imagens/continua2_F01.gif" width="24" height="24" onClick="move(document.frmCadTurma.selUnidade_Selecionado,document.frmCadTurma.selUnidade,1);submet_pagina('','Unidade');"></div></td>
						</tr>
						<tr>
						  <td height="7" align="center" valign="middle"></td>
						</tr>				
				  </table>
				   
				   </td>				 
			    </tr>
				 <tr>
				   <td height="54"></td>
				   <td height="54" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Multiplicador:</b></font></td>
				   <td height="54" valign="top" align="left">
				   	<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr>
						<td width="29%" height="30"><input name="rdMultplicador" type="radio" value="0" onClick="MostraEscondeMult();VerificaCurso('Unidade');VerificaUnidades('Unidade');">&nbsp;<font face="Verdana" color="#330099" size="2"><b>Unidade</b></font></td>
						<td width="71%">
							<select size="1" name="selMultiplic">
							  <option value="0">== Selecione o Multiplicador ==</option>
								<%
								do until rdsMultiplicadorUnid.eof = true
									  if cint(strMultiplic) = cint(rdsMultiplicadorUnid("MULT_NR_CD_ID_MULT")) then%>
										  <option value="<%=rdsMultiplicadorUnid("MULT_NR_CD_ID_MULT")%>" selected><%=rdsMultiplicadorUnid("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorUnid("MULT_NR_CD_CHAVE")%></option>
									  <%else%>
											<option value="<%=rdsMultiplicadorUnid("MULT_NR_CD_ID_MULT")%>"><%=rdsMultiplicadorUnid("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorUnid("MULT_NR_CD_CHAVE")%></option>
									  <%end if						
									rdsMultiplicadorUnid.movenext
								loop
								
								rdsMultiplicadorUnid.close
								set rdsMultiplicadorUnid = nothing						
								%>
							</select>	
							&nbsp;
						</td>
					</tr>
					<tr>
					  <td height="32"><input name="rdMultplicador" type="radio" value="1" onClick="MostraEscondeMult();VerificaCurso('Area');VerificaUnidades('Area');">&nbsp;<font face="Verdana" color="#330099" size="2"><b>Área de Negócio</b></font></td>
						<td>
							<select size="1" name="selMultiplicArea">
							  <option value="0">== Selecione o Multiplicador ==</option>
								<%
								do until rdsMultiplicadorArea.eof = true
									  if trim(strMultiplic) = trim(rdsMultiplicadorArea("MULT_NR_CD_CHAVE")) then%>
										  <option value="<%=rdsMultiplicadorArea("MULT_NR_CD_CHAVE")%>" selected><%=rdsMultiplicadorArea("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorArea("MULT_NR_CD_CHAVE")%></option>
									  <%else%>
											<option value="<%=rdsMultiplicadorArea("MULT_NR_CD_CHAVE")%>"><%=rdsMultiplicadorArea("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorArea("MULT_NR_CD_CHAVE")%></option>
									  <%end if						
									rdsMultiplicadorArea.movenext
								loop
								
								rdsMultiplicadorArea.close
								set rdsMultiplicadorArea = nothing						
								%>								
							</select>							
							&nbsp;
						</td>
					</tr>
					<tr>
						<td><input name="rdMultplicador" type="radio" value="2" checked onClick="MostraEscondeMult();VerificaCurso('Companhia');">&nbsp;<font face="Verdana" color="#330099" size="2"><b>Companhia</b></font></td>
						<td>
							<select size="1" name="selMultiplicComp">
							  <option value="0">== Selecione o Multiplicador ==</option>								
								<%
								do until rdsMultiplicadorComp.eof = true
									  if trim(strMultiplic) = trim(rdsMultiplicadorComp("MULT_NR_CD_CHAVE")) then%>
										  <option value="<%=rdsMultiplicadorComp("MULT_NR_CD_CHAVE")%>" selected><%=rdsMultiplicadorComp("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorComp("MULT_NR_CD_CHAVE")%></option>
									  <%else%>
											<option value="<%=rdsMultiplicadorComp("MULT_NR_CD_CHAVE")%>"><%=rdsMultiplicadorComp("MULT_TX_NOME_MULTIPLICADOR") & " - " & rdsMultiplicadorComp("MULT_NR_CD_CHAVE")%></option>
									  <%end if						
									rdsMultiplicadorComp.movenext
								loop
								
								rdsMultiplicadorComp.close
								set rdsMultiplicadorComp = nothing						
								%>								
							</select>			
						</td>
					  </tr>
				   </table>				   
				  </td>
			    </tr>
			    <tr> 
				  <td width="1" height="30"></td>
				  <td width="136" height="30" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mandante:</b></font></td>
				  <td height="30" valign="middle" align="left" width="596"><font face="Verdana" size="2" color="#330099"><b>
				    <input type="text" name="txtMandante" maxlength="100" size="50" value="<%=strMandante%>">
				  </b></font> 
				  </td>
				</tr> 			
				
				<tr> 
				  <td width="1" height="34"></td>
				  <td width="136" height="34" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data de Início:</b></font></td>
				  <td height="34" valign="middle" align="left" width="596"> 
					<input type="text" name="txtDtIni" maxlength="10" size="10" value="<%=strDtIni%>" onChange="javescript:CarregaDataFim(this.value);">
				    <a href="javascript:show_calendar(true,'frmCadTurma.txtDtIni','DD/MM/YYYY')"><img src="../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>				  
				 	&nbsp;&nbsp;&nbsp;<font face="Verdana" size="2" color="#330099"><b>Hora
				 	 
				 	de Início:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="text" name="txtHrIni" maxlength="5" size="7" value="<%=strHrIni%>">
				  </td>
				</tr>   
								
				<tr> 
				  <td width="1" height="34"></td>
				  <td width="136" height="34" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data de T&eacute;rmino:</b></font></td>
				  <td height="34" valign="middle" align="left" width="596"> 
					<input type="text" name="txtDtFim" maxlength="10" size="10" value="<%=strDtFim%>">
				    <a href="javascript:show_calendar(true,'frmCadTurma.txtDtFim','DD/MM/YYYY')"><img src="../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>				  
				  	&nbsp;&nbsp;&nbsp;<font face="Verdana" size="2" color="#330099"><b>Hora de Término:</b></font>
					<input type="text" name="txtHrFim" maxlength="5" size="7" value="<%=strHrFim%>">
				  </td>
				</tr>   								
		  </table>
	</form>
	
	<script language="javascript">	
		//*** CHAMADA PARA FUNÇÃO QUE DESABILITA OS COMBOS	
		MostraEscondeMult();					
	</script>
	
	</body>
	<%	
	db_banco.close
	set db_banco = nothing
	%>
</html>

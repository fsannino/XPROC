<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strAcao = trim(Request("parAcao"))
	
if trim(Request("selCorte")) <> "" then	
	Session("Corte") = trim(Request("selCorte"))	
end if		

strCdDescMultiplicador	= Ucase(trim(Request("txtCdMultiplicador")))	
strNomeMultiplicador	= strCdDescMultiplicador 'Ucase(trim(Request("txtNomeMultiplicador")))	
'strRestrViagem			= trim(Request("pRestrViagem"))]
intCdDiretoria			= trim(Request("selDiretoria"))
intCDsUnidades			= trim(Request("txtUnid_Selecionados"))
intCDsCurso				= trim(Request("txtCurso_Selecionados"))
				
strVetMult = split(trim(Request("selMultiplicador")),"|")
if trim(Request("selMultiplicador")) <> "" then			
	intCdMultiplicador = Ucase(strVetMult(0))
	intTipoMultiplicador = strVetMult(1)
'elseif trim(Request("txtCdMultiplicador")) <> "" then		
	'intCdMultiplicador = Ucase(trim(Request("txtCdMultiplicador")))
end if			
				
'Response.write "strAcao - " & strAcao & "<br>"	
'Response.write "intCdMultiplicador - " & intCdMultiplicador & "<br>"	
'Response.write "strCdDescMultiplicador - " & strCdDescMultiplicador & "<br>"
'Response.write "strNomeMultiplicador - " & strNomeMultiplicador & "<br>"
'Response.write "intCdDiretoria - " & intCdDiretoria & "<br>"
'Response.write "strRestrViagem - " & strRestrViagem & "<br>"
'Response.write "intCDsUnidades - " & intCDsUnidades & "<br>"
'Response.write "intCDsCurso - " & intCDsCurso & "<br><br>"
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Multiplicador"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Multiplicador"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Multiplicador"
elseif strAcao = "AS" then
	strNomeAcao = "Associação de Multiplicador Extra"
elseif strAcao = "APROVA" then	
	strNomeAcao = "Aprovação de Multiplicador"
end if
					
					
strMSG =  ""		
'************************************** ASSOCIAÇÃO DE MULTIPLICADOR EXTRA************************************************
if strAcao = "AS" then	

	strVetMultTrue = split(trim(Request("selMultiplicadorTrue")),"|")
	if trim(Request("selMultiplicadorTrue")) <> "" then			
		intCdMultiplicadorTrue = Ucase(strVetMultTrue(0))
		intTipoMultiplicadorTrue = strVetMultTrue(1)	
	end if			

	'Response.write "intCdMultiplicador - " & intCdMultiplicador & "<br>"
	'Response.write "intCdMultiplicadorTrue - " & intCdMultiplicadorTrue & "<br>"
	'Response.write "intTipoMultiplicadorTrue - " & intTipoMultiplicadorTrue & "<br><br><br>"
	
	SQLAltMultTurma = ""		
	SQLAltMultTurma = SQLAltMultTurma & "UPDATE GRADE_TURMA "
	SQLAltMultTurma = SQLAltMultTurma & "SET MULT_NR_CD_ID_MULT = " & intCdMultiplicadorTrue
	SQLAltMultTurma = SQLAltMultTurma & " WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	SQLAltMultTurma = SQLAltMultTurma & " AND CORT_CD_CORTE=" & Session("Corte")
	'Response.WRITE SQLAltMultTurma & "<br>"
	'response.end	
	
	SQLDELMultExtra = ""		
	SQLDELMultExtra = SQLDELMultExtra & "DELETE FROM GRADE_MULTIPLICADOR "
	SQLDELMultExtra = SQLDELMultExtra & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	SQLDELMultExtra = SQLDELMultExtra & " AND CORT_CD_CORTE = " & Session("Corte") 
	'Response.WRITE SQLDELMultExtra & "<br>"
	'response.end	
	
	SQLDELMultCursoExtra = ""		
	SQLDELMultCursoExtra = SQLDELMultCursoExtra & "DELETE FROM GRADE_MULTIPLICADOR_CURSO "
	SQLDELMultCursoExtra = SQLDELMultCursoExtra & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	SQLDELMultCursoExtra = SQLDELMultCursoExtra & " AND CORT_CD_CORTE = " & Session("Corte") 
	'Response.WRITE SQLDELMultCursoExtra & "<br>"
	'response.end	
	
	on error resume next
			db_banco.Execute(SQLAltMultTurma)		
			db_banco.Execute(SQLDELMultExtra)			
			db_banco.Execute(SQLDELMultCursoExtra)				
	
	if err.number = 0 then		
		strMSG = "Multiplicador Extra foi associado ao Multiplicador com sucesso."
	else
		'if err.number = "2147217873" then		
			strMSG = "Houve um erro na associação do Multiplicador Extra ao Multiplicador (" & err.description & " - " & err.number & ")"
		'end if
	end if			
	
	
'************************************** INCLUSÃO DE MULTIPLICADOR ************************************************
elseif strAcao = "I" then	

	intTipoMultiplicador = 3
	strTipoMultiplicador = "EXTRA"
	
	strVerificaMultiplicador = ""
	strVerificaMultiplicador = strVerificaMultiplicador & "SELECT MULT_TX_NOME_MULTIPLICADOR "
	strVerificaMultiplicador = strVerificaMultiplicador & "FROM GRADE_MULTIPLICADOR "
	strVerificaMultiplicador = strVerificaMultiplicador & "WHERE MULT_TX_NOME_MULTIPLICADOR = '" & strNomeMultiplicador & "'"	
	strVerificaMultiplicador = strVerificaMultiplicador & " AND CORT_CD_CORTE = " & Session("Corte") 	
	strVerificaMultiplicador = strVerificaMultiplicador & " AND MULT_NR_CD_ID_MULT = " & intTipoMultiplicador
	'Response.write strVerificaMultiplicador
	'Response.end
		
	Set rdsVerificaMultiplicador = db_banco.Execute(strVerificaMultiplicador)			
	
	if not rdsVerificaMultiplicador.EOF then
		strMSG = "Já existe multiplicador cadastrado com o nome - " & rdsVerificaMultiplicador("MULT_TX_NOME_MULTIPLICADOR") & "."
	else		
	
		'*** MONTA O NOVO ID PARA O NOVO REGISTRO
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(MULT_NR_CD_ID_MULT) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_MULTIPLICADOR "	
		strVerificaCod = strVerificaCod & "WHERE CORT_CD_CORTE = " & Session("Corte") 		
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCdMultiplicador = 1
			else
				intCdMultiplicador = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCdMultiplicador = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
				
		strSQLIncMultiplicador = ""
		strSQLIncMultiplicador = strSQLIncMultiplicador & "INSERT INTO GRADE_MULTIPLICADOR (CORT_CD_CORTE, MULT_NR_CD_ID_MULT, "
		strSQLIncMultiplicador = strSQLIncMultiplicador & "MULT_NR_TIPO_MULTIPLICADOR, MULT_TX_TIPO_MULTIPLICADOR, MULT_TX_NOME_MULTIPLICADOR, "
		strSQLIncMultiplicador = strSQLIncMultiplicador & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncMultiplicador = strSQLIncMultiplicador & "VALUES(" & Session("Corte") & "," & intCdMultiplicador & "," & intTipoMultiplicador & ",'" & strTipoMultiplicador & "','"
		strSQLIncMultiplicador = strSQLIncMultiplicador & strNomeMultiplicador & "','I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLIncMultiplicador & "<br><br>"
		'Response.end	
				
		on error resume next
			db_banco.Execute(strSQLIncMultiplicador)	
			
		'*** CADASTRA AS NOVOS REGISTROS NA TABELA DE GRADE_MULTIPLICADOR_ORGAO_MENOR ***			
		vetCDsUnidade = split(intCDsUnidades,",")
					
		p = 0			
		for p = lbound(vetCDsUnidade) to Ubound(vetCDsUnidade)				
						
			if vetCDsUnidade(p) <> "" then				
			
				intCDUnidResult = cstr(vetCDsUnidade(p))	

				strSQLOrgMenor = ""
				strSQLOrgMenor = strSQLOrgMenor & "SELECT CORT_CD_CORTE, UNID_CD_UNIDADE, ORME_CD_ORG_MENOR "
				strSQLOrgMenor = strSQLOrgMenor & "FROM GRADE_UNIDADE_ORGAO_MENOR "
				strSQLOrgMenor = strSQLOrgMenor & "WHERE CORT_CD_CORTE = " & Session("Corte") 				
				strSQLOrgMenor = strSQLOrgMenor & " AND UNID_CD_UNIDADE = " & intCDUnidResult
				'Response.write strSQLOrgMenor & "<br><br>"
				'Response.end
				
				set rsOrgMenor = db_banco.Execute(strSQLOrgMenor)
				
				do while not rsOrgMenor.eof
				
					strNumOrgMenor = rsOrgMenor("ORME_CD_ORG_MENOR")
													
					strSQLIncMultiplicUnid = ""		
					strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "INSERT INTO GRADE_MULTIPLICADOR_ORGAO_MENOR "
					strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "(CORT_CD_CORTE, MULT_NR_CD_ID_MULT, ORME_CD_ORG_MENOR, "
					strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
					strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "VALUES(" & Session("Corte") & "," & intCdMultiplicador & ",'" & strNumOrgMenor & "'," 
					strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "'I','" & Session("CdUsuario") & "',GETDATE())"					
					'response.write strSQLIncMultiplicUnid & "<br><br>"
					'Response.end					
					db_banco.Execute(strSQLIncMultiplicUnid)
					
					rsOrgMenor.movenext
				loop
				
				rsOrgMenor.close
				set rsOrgMenor = nothing				
			end if						
		next			
				
		'*** CADASTRA AS NOVOS REGISTROS NA TABELA DE GRADE_MULTIPLICADOR_CURSO ***			
		vetCDsCurso = split(intCDsCurso,",")
					
		r = 0			
		for r = lbound(vetCDsCurso) to Ubound(vetCDsCurso)				
						
			if vetCDsCurso(r) <> "" then				
			
				intCDCursoResult = cstr(vetCDsCurso(r))
						
				strSQLIncMultiplicCurso = ""		
				strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "INSERT INTO GRADE_MULTIPLICADOR_CURSO "
				strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "(CORT_CD_CORTE, MULT_NR_CD_ID_MULT, CURS_CD_CURSO, "
				strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
				strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "VALUES(" & Session("Corte") & "," & intCdMultiplicador & ",'" & intCDCursoResult & "'," 
				strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "'I','" & Session("CdUsuario") & "',GETDATE())"					
				'response.write strSQLIncMultiplicCurso & "<br><br>"
				'Response.end					
				db_banco.Execute(strSQLIncMultiplicCurso)
			end if
		next		
			
		if err.number = 0 then		
			strMSG = "Multiplicador foi incluído com sucesso."
		else
			'if err.number = "2147217873" then		
				strMSG = "Houve um erro na inclusão do multiplicador (" & err.description & " - " & err.number & ")"
			'end if
		end if	
	end if
	
	rdsVerificaMultiplicador.close
	set rdsVerificaMultiplicador = nothing
	
	
'************************************** ALTERAÇÃO DE MULTIPLICADOR ************************************************	
elseif strAcao = "A" then			
				
	intTipoMultiplicador = trim(Request("pintTipoMult"))	
	strTipoMultiplicador = trim(Request("pstrTipoMult"))	
		
	'****** MULTIPLICADOR MULTIPLICADOR ORGAO MENOR *********
		
	'*** LIMPA OS REGISTROS DA TABELA GRADE_MULTIPLICADOR_ORGAO_MENOR **
	strSQLDelMultiplicOrgMenor = ""
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & "DELETE FROM GRADE_MULTIPLICADOR_ORGAO_MENOR "	
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & " AND CORT_CD_CORTE=" & Session("Corte")
	'response.write strSQLDelMultiplicOrgMenor
	'Response.end		
	
	db_banco.Execute(strSQLDelMultiplicOrgMenor)	
	
	'*** CADASTRA AS NOVOS REGISTROS NA TABELA DE GRADE_MULTIPLICADOR_ORGAO_MENOR ***			
	vetCDsUnidade = split(intCDsUnidades,",")
				
	p = 0			
	for p = lbound(vetCDsUnidade) to Ubound(vetCDsUnidade)				
					
		if vetCDsUnidade(p) <> "" then				
		
			intCDUnidResult = cstr(vetCDsUnidade(p))	

			strSQLOrgMenor = ""
			strSQLOrgMenor = strSQLOrgMenor & "SELECT CORT_CD_CORTE, UNID_CD_UNIDADE, ORME_CD_ORG_MENOR "
			strSQLOrgMenor = strSQLOrgMenor & "FROM GRADE_UNIDADE_ORGAO_MENOR "
			strSQLOrgMenor = strSQLOrgMenor & "WHERE CORT_CD_CORTE = " & Session("Corte") 				
			strSQLOrgMenor = strSQLOrgMenor & " AND UNID_CD_UNIDADE = " & intCDUnidResult
			'Response.write strSQLOrgMenor & "<br><br>"
			'Response.end
			
			set rsOrgMenor = db_banco.Execute(strSQLOrgMenor)
			
			do while not rsOrgMenor.eof
			
				strNumOrgMenor = rsOrgMenor("ORME_CD_ORG_MENOR")
												
				strSQLIncMultiplicUnid = ""		
				strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "INSERT INTO GRADE_MULTIPLICADOR_ORGAO_MENOR "
				strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "(CORT_CD_CORTE, MULT_NR_CD_ID_MULT, ORME_CD_ORG_MENOR, "
				strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
				strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "VALUES(" & Session("Corte") & "," & intCdMultiplicador & ",'" & strNumOrgMenor & "'," 
				strSQLIncMultiplicUnid = strSQLIncMultiplicUnid & "'I','" & Session("CdUsuario") & "',GETDATE())"					
				'response.write strSQLIncMultiplicUnid & "<br><br>"
				'Response.end					
				db_banco.Execute(strSQLIncMultiplicUnid)
				
				rsOrgMenor.movenext
			loop
			
			rsOrgMenor.close
			set rsOrgMenor = nothing				
		end if						
	next			
	
	'****** MULTIPLICADOR CURSO *********
					
	'*** LIMPA OS REGISTROS DA TABELA GRADE_MULTIPLICADOR_CURSO ***
	strSQLDelMultiplicCurso = ""
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & "DELETE FROM GRADE_MULTIPLICADOR_CURSO "	
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & " AND CORT_CD_CORTE=" & Session("Corte")
	'response.write strSQLDelMultiplicCurso
	'Response.end		
	
	db_banco.Execute(strSQLDelMultiplicCurso)					
					
	'*** CADASTRA AS NOVOS REGISTROS NA TABELA DE GRADE_MULTIPLICADOR_CURSO ***		
	vetCDsCurso = split(intCDsCurso,",")
				
	r = 0			
	for r = lbound(vetCDsCurso) to Ubound(vetCDsCurso)				
				
		if vetCDsCurso(r) <> "" then				
			intCDCursoResult = vetCDsCurso(r)
					
			strSQLIncMultiplicCurso = ""		
			strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "INSERT INTO GRADE_MULTIPLICADOR_CURSO "
			strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "(CORT_CD_CORTE, MULT_NR_CD_ID_MULT, CURS_CD_CURSO, "
			strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
			strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "VALUES(" & Session("Corte") & ",'" & intCdMultiplicador & "','" & intCDCursoResult & "'," 
			strSQLIncMultiplicCurso = strSQLIncMultiplicCurso & "'I','" & Session("CdUsuario") & "',GETDATE())"					
			'response.write strSQLIncMultiplicCurso & "<br><br>"
			'Response.end					
			db_banco.Execute(strSQLIncMultiplicCurso)
		end if
	next						
		
	strSQLAltFeriado = ""
	strSQLAltFeriado = strSQLAltFeriado & "UPDATE GRADE_MULTIPLICADOR "
	strSQLAltFeriado = strSQLAltFeriado & "SET MULT_TX_NOME_MULTIPLICADOR = '" & strNomeMultiplicador & "',"	
	'strSQLAltFeriado = strSQLAltFeriado & "ORME_CD_ORG_MENOR = " & intCdDiretoria & ","			
	'strSQLAltFeriado = strSQLAltFeriado & "MULT_TX_TIPO_MULTIPLICADOR = '" & strTipoMultiplicador & "',"		
	'strSQLAltFeriado = strSQLAltFeriado & "MULT_TX_RESTRICAO_VIAGEM = '" & strRestrViagem & "',"		
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltFeriado = strSQLAltFeriado & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLAltFeriado = strSQLAltFeriado & " AND CORT_CD_CORTE=" & Session("Corte")
	strSQLAltFeriado = strSQLAltFeriado & " AND MULT_NR_TIPO_MULTIPLICADOR=" & intTipoMultiplicador
	'Response.write strSQLAltFeriado
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltFeriado)
					
	if err.number = 0 then		
		strMSG = "Multiplicador foi alterado com sucesso."
	else
		strMSG = "Houve um erro na alteração do multiplicador (" & err.description & ")"
	end if			
	
	
'************************************** APROVAÇÃO DE MULTIPLICADOR ************************************************	
elseif strAcao = "APROVA" then 
		
	strSQLMultiplicadorAprov = Request("txtQuery")
	strCursoAprov = Request("hdStrCurso")
	
	'*** PEGA O NOME DO CORTE ***				 
	strSQLCorte = ""
	strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
	strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
	strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
	
	set rsCorte = db_banco.Execute(strSQLCorte)		
	
	strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")
	
	rsCorte.close
	set rsCorte = nothing		
	
	'*** SELECIONA OS MULTIPLICADORES PARA O CURSO SELECIONADO
	set rsMultiplicadorAprov = db_banco.execute(strSQLMultiplicadorAprov)
		
	on error resume next
		i = 0
		st_atual = 0
		tot_reg = rsMultiplicadorAprov.RecordCount
				
		do until i = tot_reg
							
			st_atual = cint(request(trim(rsMultiplicadorAprov("MULT_NR_CD_ID_MULT")) & "_" & trim(rsMultiplicadorAprov("CURS_CD_CURSO"))))
													
			if st_atual = 1 then				
				sqlAprov = ""
				sqlAprov = sqlAprov & "UPDATE GRADE_MULTIPLICADOR_CURSO "				
				sqlAprov = sqlAprov & "SET MULT_TX_APROVEITAMENTO = GETDATE() "
				sqlAprov = sqlAprov & "WHERE CURS_CD_CURSO = '" & strCursoAprov & "' "
				sqlAprov = sqlAprov & "AND CORT_CD_CORTE = " & Session("Corte")
				sqlAprov = sqlAprov & "AND MULT_NR_CD_ID_MULT = " & rsMultiplicadorAprov("MULT_NR_CD_ID_MULT")				
				'Response.write "Entrou no if" & sqlAprov & "<br>"		
				'Response.end						
				db_banco.execute(sqlAprov)
			else				
				sqlAprov = ""
				sqlAprov = sqlAprov & "UPDATE GRADE_MULTIPLICADOR_CURSO "
				sqlAprov = sqlAprov & "SET MULT_TX_APROVEITAMENTO = NULL "
				sqlAprov = sqlAprov & "WHERE CURS_CD_CURSO = '" & strCursoAprov & "' "
				sqlAprov = sqlAprov & "AND CORT_CD_CORTE = " & Session("Corte")
				sqlAprov = sqlAprov & "AND MULT_NR_CD_ID_MULT = " & rsMultiplicadorAprov("MULT_NR_CD_ID_MULT")
				'Response.write "Entrou no else" & "<br>"					
				db_banco.execute(sqlAprov)
			end if
			
			i = i + 1
			rsMultiplicadorAprov.movenext			
		loop
		
	if err.number = 0 then		
		strMSG = "Multiplicador foi aprovado para o Curso - " & strCursoAprov & " , para o corte - " & strNomeCorte & " com sucesso."
	else
		strMSG = "Houve um erro na aprovado do multiplicador (" & err.description & ")"
	end if	
	
'************************************** EXCLUSÃO DE MULTIPLICADOR ************************************************	
elseif strAcao = "E" then 
				
	intTipoMultiplicador = trim(Request("pintTipoMult"))	
	strTipoMultiplicador = trim(Request("pstrTipoMult"))				
	
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmMult = ""	
	
	strSQLVeriMultTurma = ""			
	strSQLVeriMultTurma = strSQLVeriMultTurma & "SELECT DISTINCT MULT.MULT_TX_NOME_MULTIPLICADOR "
	strSQLVeriMultTurma = strSQLVeriMultTurma & "FROM GRADE_TURMA TURMA, GRADE_MULTIPLICADOR MULT "
	strSQLVeriMultTurma = strSQLVeriMultTurma & "WHERE TURMA.MULT_NR_CD_ID_MULT = MULT.MULT_NR_CD_ID_MULT "
	strSQLVeriMultTurma = strSQLVeriMultTurma & "AND MULT.MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLVeriMultTurma = strSQLVeriMultTurma & " AND TURMA.CORT_CD_CORTE = " & Session("Corte") 	
	strSQLVeriMultTurma = strSQLVeriMultTurma & " AND MULT.CORT_CD_CORTE = " & Session("Corte") 	
	'Response.write strSQLVeriMultTurma & "<br>"
	'Response.end		
			
	set rsVeriMultTurma = db_banco.Execute(strSQLVeriMultTurma)		
			
	if not rsVeriMultTurma.eof then	
		blnPodeDeletar = False
		strAssociacao = "Turma"
		strNmMult = rsVeriMultTurma("MULT_TX_NOME_MULTIPLICADOR")
	end if
	
	rsVeriMultTurma.close
	set rsVeriMultTurma = nothing
	
	'*** LIMPA OS REGISTROS DA TABELA GRADE_MULTIPLICADOR_ORGAO_MENOR **
	strSQLDelMultiplicOrgMenor = ""
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & "DELETE FROM GRADE_MULTIPLICADOR_ORGAO_MENOR "	
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLDelMultiplicOrgMenor = strSQLDelMultiplicOrgMenor & " AND CORT_CD_CORTE=" & Session("Corte")
	'response.write strSQLDelMultiplicOrgMenor
	'Response.end		
				
	'*** LIMPA OS REGISTROS DA TABELA PARA O GRADE_MULTIPLICADOR_CURSO **
	strSQLDelMultiplicCurso = ""
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & "DELETE FROM GRADE_MULTIPLICADOR_CURSO "	
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & "WHERE MULT_NR_CD_ID_MULT = '" & intCdMultiplicador & "' "
	strSQLDelMultiplicCurso = strSQLDelMultiplicCurso & "AND CORT_CD_CORTE=" & Session("Corte")
	'response.write strSQLDelMultiplicCurso
	'Response.end		
		
	'*** LIMPA OS REGISTROS DA TABELA PARA O GRADE_MULTIPLICADOR **
	strSQLDelFeriado = ""
	strSQLDelFeriado = strSQLDelFeriado & "DELETE FROM GRADE_MULTIPLICADOR "	
	strSQLDelFeriado = strSQLDelFeriado & "WHERE MULT_NR_CD_ID_MULT = '" & intCdMultiplicador & "' "
	strSQLDelFeriado = strSQLDelFeriado & "AND CORT_CD_CORTE=" & Session("Corte")
	'strSQLDelFeriado = strSQLDelFeriado & " AND MULT_NR_TIPO_MULTIPLICADOR=" & intTipoMultiplicador
	'Response.write strSQLDelFeriado
	'Response.end
	
	on error resume next
		
		if blnPodeDeletar then			
		
			db_banco.Execute(strSQLDelMultiplicOrgMenor)		
			db_banco.Execute(strSQLDelMultiplicCurso)		
			db_banco.Execute(strSQLDelFeriado)
		
			if err.number = 0 then		
				strMSG = "Multiplicador foi excluído com sucesso."
			else
				strMSG = "Houve um erro na exclusão do multiplicador (" & err.description & ")"
			end if	
		
		else
			strMSG = "O Multiplicador - " & strNmMult & " não pode ser excluído, pois existe associação do mesmo em " & strAssociacao & "."
		end if
end if
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" action="valida_cad_curso.asp" name="frm1">	
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
				  <table width="625" border="0" align="center">
					<tr>
						<td width="26"></td>
					  <td width="50"></td>
					  <td width="26"></td>
					  <td width="195"></td>
						<td width="27"></td>  <td width="50"></td>
					  <td width="28"></td>
					  <td width="26">&nbsp;</td>
					  <td width="159"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
					
			  <table width="847" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td width="845" height="15">
				  </td>
				</tr>
				<tr>
				  <td width="845">				  
					<div align="center"><font face="Verdana" color="#330099" size="3"><b><%=strNomeAcao%> - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td width="845">&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="849" height="142">
					  <tr>
						
				  <td width="205" height="29"></td>
						
				  <td width="93" height="29" valign="middle" align="left"></td>
						
				  <td height="29" valign="middle" align="left" colspan="2"> 				 
				  	<b><font face="Verdana" color="#330099" size="2"><%=strMSG%></font></b> 
				  </td>
						
					  </tr>
				  
					  <tr>
						
				  <td width="205" height="1"></td>
						
				  <td width="93" height="1" valign="middle" align="left"></td>
						
				  <td height="1" valign="middle" align="left" colspan="2"> 
				  </td>
					  </tr>
					  <tr>
						
				  <td width="205" height="1"></td>
						
				  <td width="93" height="1" valign="middle" align="left"></td>
						
				  <td height="1" valign="middle" align="left" colspan="2"> 
				  </td>
					  </tr>
					  <tr>
						
				  <td width="205" height="1"></td>
						
				  <td width="93" height="1" valign="middle" align="left"></td>
						
				  <td height="1" valign="middle" align="left" colspan="2"> 
				  </td>
					  </tr>
					  <tr>
					    <td height="1"></td>
					    <td height="1" valign="middle" align="left"></td>
					    <td height="1" valign="middle" align="left">&nbsp;</td>
					    <td height="1" valign="middle" align="left">&nbsp;</td>
			    </tr>
					  <tr>
						
				  <td width="205" height="35"></td>
						
				  <td width="93" height="35" valign="middle" align="left"></td>
						
				  <td height="35" valign="middle" align="left" width="29"> 
					<a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"> 
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="35" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
				</tr>
					
				<tr>						
				  <td width="205" height="29"></td>						
				  <td width="93" height="29" valign="middle" align="left"></td>						
				  <td height="29" valign="middle" align="left" width="29"> 				 
				    <%
					if strAcao = "APROVA" then	
					%>
						<a href="sel_multiplicador_aprov.asp?selCorte=<%=Session("Corte")%>">		
					<%
					else
					%>	
						<a href="sel_multiplicador.asp?selCorte=<%=Session("Corte")%>">		
					<%
					end if
					%>   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Multiplicador</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
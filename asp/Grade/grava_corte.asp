<%
server.ScriptTimeout=99999999
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao 		= trim(Request("parAcao"))
strNomeCorte	= Ucase(trim(Request("txtNomeCorte")))		

if trim(Request("txtDtCorte")) <> "" then
strDtCorte = MontaDataHora(trim(Request("txtDtCorte")),4)
else
strDtCorte = ""
end if
	
'Response.write "strAcao - " & strAcao & "<br>"
'Response.write "strCdCorte - " & strCdCorte & "<br>"
'Response.write "strNomeCorte - " & strNomeCorte & "<br>"
'Response.write "strDtCorte - " & strDtCorte & "<br>"	
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Corte"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Corte"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Corte"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE CORTE ************************************************
if strAcao = "I" then	

	strVerificaCorte = ""
	strVerificaCorte = strVerificaCorte & "SELECT CORT_TX_DESC_CORTE "
	strVerificaCorte = strVerificaCorte & "FROM GRADE_CORTE "
	strVerificaCorte = strVerificaCorte & "WHERE CORT_TX_DESC_CORTE = '" & strNomeCorte & "'"		
	'Response.write strVerificaCorte
	'Response.end
		
	Set rdsVerificaCorte = db_banco.Execute(strVerificaCorte)			
	
	if not rdsVerificaCorte.EOF then
		strMSG = "Já existe corte cadastrado com o nome - " & rdsVerificaCorte("CORT_TX_DESC_CORTE") & "."
	else			
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(CORT_CD_CORTE) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_CORTE "			
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCdCorte = 1
			else				
				intCdCorte = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCdCorte = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
	
		strSQLIncCorte = ""
		strSQLIncCorte = strSQLIncCorte & "INSERT INTO GRADE_CORTE (CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE, "
		strSQLIncCorte = strSQLIncCorte & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncCorte = strSQLIncCorte & "VALUES(" & intCdCorte & ",'" & strNomeCorte & "','" & strDtCorte & "',"
		strSQLIncCorte = strSQLIncCorte & "'I','" & Session("CdUsuario") & "',GETDATE())"	
	
		db_banco.Execute(strSQLIncCorte)	

		if err.number = 0 then		
			strMSG = "Corte foi incluído com sucesso."
		else
			strMSG = "Houve um erro na inclusão do corte (" & err.description & ")"
		end if	
	end if
	
	rdsVerificaCorte.close
	set rdsVerificaCorte = nothing
	
	Corte_Atual = intCdCorte
	Corte_Anterior = Corte_Atual - 1
	
	if Corte_Anterior < 0 then
		Corte_Anterior = 0
	end if
	
	err.clear
	
	'=============================== COPIA DADOS DE CORTE ANTERIOR PARA CORTE ATUAL ====================

	'======== COPIA CENTROS DE TREINAMENTO DO CORTE ANTERIOR ===============	
	
	ssql=""
	ssql="INSERT INTO GRADE_CENTRO_TREINAMENTO("
	ssql=ssql+"CORT_CD_CORTE, "
	ssql=ssql+"CTRO_CD_CENTRO_TREINAMENTO,"
	ssql=ssql+"LOC_CD_LOCALIDADE,"
	ssql=ssql+"CTRO_TX_NOME_CENTRO_TREINAMENTO,"
	ssql=ssql+"ATUA_TX_OPERACAO,"
	ssql=ssql+"ATUA_CD_NR_USUARIO,"
	ssql=ssql+"ATUA_DT_ATUALIZACAO) "
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+"CTRO_CD_CENTRO_TREINAMENTO,"
	ssql=ssql+"LOC_CD_LOCALIDADE,"
	ssql=ssql+"CTRO_TX_NOME_CENTRO_TREINAMENTO,"
	ssql=ssql+"ATUA_TX_OPERACAO,"
	ssql=ssql+"ATUA_CD_NR_USUARIO,"
	ssql=ssql+"ATUA_DT_ATUALIZACAO "
	ssql=ssql+"FROM GRADE_CENTRO_TREINAMENTO "
	ssql=ssql+"WHERE CORT_CD_CORTE=" & Corte_Anterior & ")" 
	
	db_banco.Execute(ssql)	
	
	'======== COPIA SALAS DO CORTE ANTERIOR ===============	
	
	ssql=""
	ssql="INSERT INTO GRADE_SALA("
	ssql=ssql+"CORT_CD_CORTE,"
	ssql=ssql+" SALA_CD_SALA,"
	ssql=ssql+" CTRO_CD_CENTRO_TREINAMENTO,"
	ssql=ssql+" SALA_TX_NOME_SALA,"
	ssql=ssql+" SALA_NUM_CAPACIDADE,"
	ssql=ssql+" SALA_CD_UC,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO)"
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+" SALA_CD_SALA,"
	ssql=ssql+" CTRO_CD_CENTRO_TREINAMENTO,"
	ssql=ssql+" SALA_TX_NOME_SALA,"
	ssql=ssql+" SALA_NUM_CAPACIDADE,"
	ssql=ssql+" SALA_CD_UC,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO"
	ssql=ssql+" FROM GRADE_SALA "
	ssql=ssql+"WHERE CORT_CD_CORTE=" & Corte_Anterior & ")" 
	
	db_banco.Execute(ssql)	
	
	'======== COPIA FERIADOS REFERENTES À SALAS DO CORTE ANTERIOR ===============	
	
	ssql=""
	ssql="INSERT INTO GRADE_FERIADO_SALA("
	ssql=ssql+"CORT_CD_CORTE,"
	ssql=ssql+" SALA_CD_SALA,"
	ssql=ssql+" FERI_CD_FERIADO,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO)"
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+" SALA_CD_SALA,"
	ssql=ssql+" FERI_CD_FERIADO,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO"
	ssql=ssql+" FROM GRADE_FERIADO_SALA "
	ssql=ssql+"WHERE CORT_CD_CORTE=" & Corte_Anterior & ")" 
	
	db_banco.Execute(ssql)

	'======== COPIA CURSOS NO CORTE ATUAL===============

	ssql=""
	ssql="INSERT INTO GRADE_CURSO("
	ssql=ssql+"CORT_CD_CORTE, "
	ssql=ssql+"CURS_CD_CURSO, "
	ssql=ssql+"CURS_TX_NOME_CURSO, "
	ssql=ssql+"CURS_NUM_CARGA_CURSO, "
	ssql=ssql+"CURS_TX_METODO_CURSO, "
	ssql=ssql+"MEPR_CD_MEGA_PROCESSO, "
	ssql=ssql+"CURS_TX_CENTRALIZADO, "
	ssql=ssql+"CURS_DT_FIM_MATERIAL_DIDATICO, "
	ssql=ssql+"CURS_TX_IN_LOCO, "	
	ssql=ssql+"ATUA_TX_OPERACAO, "
	ssql=ssql+"ATUA_CD_NR_USUARIO, "
	ssql=ssql+"ATUA_DT_ATUALIZACAO, "
	ssql=ssql+"ONDA_CD_ONDA_ABRANGENCIA)"
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+"CURS_CD_CURSO, "
	ssql=ssql+"CURS_TX_NOME_CURSO, "
	ssql=ssql+"CURS_NUM_CARGA_CURSO, "
	ssql=ssql+"CURS_TX_METODO_CURSO, "
	ssql=ssql+"MEPR_CD_MEGA_PROCESSO, "
	ssql=ssql+"CURS_TX_CENTRALIZADO, "
	ssql=ssql+"CURS_DT_FIM_MATERIAL_DIDATICO, "
	ssql=ssql+"CURS_TX_IN_LOCO, "
	ssql=ssql+"ATUA_TX_OPERACAO, "
	ssql=ssql+"ATUA_CD_NR_USUARIO, "
	ssql=ssql+"ATUA_DT_ATUALIZACAO, "
	ssql=ssql+"ONDA_CD_ONDA "	
	ssql=ssql+"FROM CURSO WHERE CURS_TX_STATUS_CURSO='1' AND ONDA_CD_ONDA IN(6,9))"
	
	db_banco.Execute(ssql)

	'======== COPIA CURSOS PRE-REQUISITOS NO CORTE ATUAL===============	
	
	ssql=""
	ssql="INSERT INTO GRADE_CURSO_PRE_REQUISITO("
	ssql=ssql+"CORT_CD_CORTE, "
	ssql=ssql+"CURS_CD_CURSO, "
	ssql=ssql+"CURS_PRE_REQUISITO, "	
	ssql=ssql+"ATUA_TX_OPERACAO, "
	ssql=ssql+"ATUA_CD_NR_USUARIO, "
	ssql=ssql+"ATUA_DT_ATUALIZACAO) "
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+"CURS_CD_CURSO, "
	ssql=ssql+"CURS_PRE_REQUISITO, "
	ssql=ssql+"ATUA_TX_OPERACAO, "
	ssql=ssql+"ATUA_CD_NR_USUARIO, "
	ssql=ssql+"ATUA_DT_ATUALIZACAO "
	ssql=ssql+"FROM CURSO_PRE_REQUISITO)"
	
	db_banco.Execute(ssql)

	'======== COPIA ORGAOS MAIOR E MENOR NO CORTE ATUAL ===============	
	
    str_SQL = ""
    str_SQL = str_SQL & " SELECT "
    str_SQL = str_SQL & " ORLO_CD_ORG_LOT "
    str_SQL = str_SQL & " , MAX(ORLO_NR_ORDEM) AS Expr1"
    str_SQL = str_SQL & " FROM dbo.ORGAO_MAIOR"
    str_SQL = str_SQL & " GROUP BY ORLO_CD_ORG_LOT"
    str_SQL = str_SQL & " ORDER BY ORLO_CD_ORG_LOT, MAX(ORLO_NR_ORDEM)"
    
    Set rst_Orgao_Maior = db_banco.Execute(str_SQL)
    
    Do While Not rst_Orgao_Maior.EOF
    
        str_SQL = ""
        str_SQL = str_SQL & " SELECT "
        str_SQL = str_SQL & " ORLO_CD_ORG_LOT"
        str_SQL = str_SQL & " , ORLO_NR_ORDEM"
        str_SQL = str_SQL & " , ORLO_SG_ORG_LOT"
        str_SQL = str_SQL & " , CILO_CD_CIDADE"
        str_SQL = str_SQL & " , ESTE_CD_UF"
        str_SQL = str_SQL & " , AGLU_CD_AGLUTINADO"
        str_SQL = str_SQL & " , ORLO_CD_GABINETE "
        str_SQL = str_SQL & " , ORLO_CD_STATUS"
        str_SQL = str_SQL & " , ORLO_NM_ORG_LOT"
        str_SQL = str_SQL & " , ORLO_IN_BLOQUEADO"
        str_SQL = str_SQL & " , ORLO_IN_BLOQUEADO_PERFIL"
        str_SQL = str_SQL & " FROM dbo.ORGAO_MAIOR"
        str_SQL = str_SQL & " WHERE "
        str_SQL = str_SQL & " ORLO_CD_ORG_LOT = " & rst_Orgao_Maior("ORLO_CD_ORG_LOT")
        str_SQL = str_SQL & " And ORLO_NR_ORDEM = " & rst_Orgao_Maior("Expr1")
        
        Set rst_Orgao_Maior_2 = db_banco.Execute(str_SQL)
        
        If Not rst_Orgao_Maior_2.EOF Then
        
            str_SQL = ""
            str_SQL = str_SQL & " INSERT INTO GRADE_ORGAO_MAIOR(  "
            str_SQL = str_SQL & " CORT_CD_CORTE"
            str_SQL = str_SQL & " ,ORLO_CD_ORG_LOT"
            str_SQL = str_SQL & " ,ORLO_SG_ORG_LOT"
            str_SQL = str_SQL & " ,CILO_CD_CIDADE"
            str_SQL = str_SQL & " ,ESTE_CD_UF"
            str_SQL = str_SQL & " ,AGLU_CD_AGLUTINADO"
            str_SQL = str_SQL & " ,ORLO_CD_GABINETE"
            str_SQL = str_SQL & " ,ORLO_CD_STATUS"
            str_SQL = str_SQL & " ,ORLO_NM_ORG_LOT"
            str_SQL = str_SQL & " ,ORLO_IN_BLOQUEADO"
            str_SQL = str_SQL & " ,ORLO_IN_BLOQUEADO_PERFIL"
            str_SQL = str_SQL & " ) values ("
            str_SQL = str_SQL & "" & Corte_Atual & ""
            str_SQL = str_SQL & " ," & rst_Orgao_Maior_2("ORLO_CD_ORG_LOT")
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ORLO_SG_ORG_LOT") & "'"
            str_SQL = str_SQL & " ," & rst_Orgao_Maior_2("CILO_CD_CIDADE")
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ESTE_CD_UF") & "'"
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("AGLU_CD_AGLUTINADO") & "'"
            str_SQL = str_SQL & " ," & rst_Orgao_Maior_2("ORLO_CD_GABINETE")
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ORLO_CD_STATUS") & "'"
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ORLO_NM_ORG_LOT") & "'"
            str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ORLO_IN_BLOQUEADO") & "'"
            If Not IsNull(rst_Orgao_Maior_2("ORLO_IN_BLOQUEADO_PERFIL")) Then
                str_SQL = str_SQL & " ,'" & rst_Orgao_Maior_2("ORLO_IN_BLOQUEADO_PERFIL") & "'"
            Else
                str_SQL = str_SQL & " ,Null"
            End If
            str_SQL = str_SQL & " )"
            
            db_banco.Execute (str_SQL)
                
            str_SQL = ""
            str_SQL = str_SQL & " Select * "
            str_SQL = str_SQL & " FROM ORGAO_MENOR"
            str_SQL = str_SQL & " WHERE ORME_CD_STATUS = 'A'"
            str_SQL = str_SQL & " AND ORLO_CD_ORG_LOT = " & rst_Orgao_Maior_2("ORLO_CD_ORG_LOT")
            str_SQL = str_SQL & " and ORLO_NR_ORDEM = " & rst_Orgao_Maior_2("ORLO_NR_ORDEM")

            Set rst_Orgao_Menor = db_banco.Execute(str_SQL)
            
            If Not rst_Orgao_Menor.EOF Then
            
            Do While Not rst_Orgao_Menor.EOF
                    
                    str_SQL = ""
                    str_SQL = str_SQL & " INSERT INTO GRADE_ORGAO_MENOR  ("
                    str_SQL = str_SQL & " CORT_CD_CORTE"
                    str_SQL = str_SQL & " ,ORME_CD_ORG_MENOR"
                    str_SQL = str_SQL & " ,ORME_SG_ORG_MENOR"
                    str_SQL = str_SQL & " ,AGLU_CD_AGLUTINADO"
                    str_SQL = str_SQL & " ,ORLO_CD_ORG_LOT"
                    str_SQL = str_SQL & " ,ORME_CD_DIVISAO"
                    str_SQL = str_SQL & " ,ORME_CD_SETOR"
                    str_SQL = str_SQL & " ,ORME_CD_SECAO"
                    str_SQL = str_SQL & " ,ORPA_CD_ORG_PAG"
                    str_SQL = str_SQL & " ,ORGM_CD_ORG_MEDICO"
                    str_SQL = str_SQL & " ,ORME_CD_STATUS"
                    str_SQL = str_SQL & " ,ESTE_CD_UF"
                    str_SQL = str_SQL & " ,ORME_NM_ORG_MENOR"
                    str_SQL = str_SQL & " ,ARTR_CD_AREA_TRAB"
                    str_SQL = str_SQL & " ,ATTR_CD_ATIVIDADE"
                    str_SQL = str_SQL & " ,CILO_NM_CIDADE"
                    str_SQL = str_SQL & " ,ORME_IN_BLOQUEADO"
                    str_SQL = str_SQL & " ,ORLO_NR_ORDEM"
                    str_SQL = str_SQL & " ,ORME_IN_BLOQUEADO_PERFIL"
                    str_SQL = str_SQL & " ) Values ("
                    str_SQL = str_SQL & "" & Corte_Atual & ""
                    str_SQL = str_SQL & " ," & right("0000000" & rst_Orgao_Menor("ORME_CD_ORG_MENOR"), 15)
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ORME_SG_ORG_MENOR") & "'"
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("AGLU_CD_AGLUTINADO")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORLO_CD_ORG_LOT")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORME_CD_DIVISAO")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORME_CD_SETOR")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORME_CD_SECAO")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORPA_CD_ORG_PAG")
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORGM_CD_ORG_MEDICO")
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ORME_CD_STATUS") & "'"
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ESTE_CD_UF") & "'"
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ORME_NM_ORG_MENOR") & "'"
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ARTR_CD_AREA_TRAB") & "'"
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ATTR_CD_ATIVIDADE") & "'"
                    If Not IsNull(rst_Orgao_Menor("CILO_NM_CIDADE")) Then
                        str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("CILO_NM_CIDADE") & "'"
                    Else
                        str_SQL = str_SQL & " ,Null "
                    End If
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ORME_IN_BLOQUEADO") & "'"
                    str_SQL = str_SQL & " ," & rst_Orgao_Menor("ORLO_NR_ORDEM")
                    str_SQL = str_SQL & " ,'" & rst_Orgao_Menor("ORME_IN_BLOQUEADO_PERFIL") & "'"
                    str_SQL = str_SQL & " )"
                    
                    db_banco.Execute (str_SQL)
                                        

 					rst_Orgao_Menor.MoveNext
                Loop
            End If
            
        End If
        
        rst_Orgao_Maior.MoveNext
    Loop	

	'======== COPIA MULTIPLICADORES NO CORTE ATUAL =============== OK!!!!!	
	
	ssql=""
	ssql="INSERT INTO GRADE_MULTIPLICADOR("
	ssql=ssql+"CORT_CD_CORTE,"
	ssql=ssql+" MULT_NR_CD_ID_MULT,"
	ssql=ssql+" MULT_NR_TIPO_MULTIPLICADOR,"
	ssql=ssql+" MULT_TX_TIPO_MULTIPLICADOR,"
	ssql=ssql+" MULT_NR_CD_CHAVE,"
	ssql=ssql+" MULT_TX_NOME_MULTIPLICADOR,"
	ssql=ssql+" ORME_CD_ORG_MENOR,"
	ssql=ssql+" APLO_NR_RELACAO_EMPREGO,"
	ssql=ssql+" MULT_TX_RESTRICAO_VIAGEM,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO)"	
	ssql=ssql+"(SELECT "
	ssql=ssql+"" & Corte_Atual & ", "
	ssql=ssql+" MULT_NR_CD_ID_MULT,"
	ssql=ssql+" MULT_NR_TIPO_MULTIPLICADOR,"
	ssql=ssql+" MULT_TX_TIPO_MULTIPLICADOR,"
	ssql=ssql+" MULT_NR_CD_CHAVE,"
	ssql=ssql+" MULT_TX_NOME_MULTIPLICADOR,"
	ssql=ssql+" ORME_CD_ORG_MENOR,"
	ssql=ssql+" APLO_NR_RELACAO_EMPREGO,"
	ssql=ssql+" MULT_TX_RESTRICAO_VIAGEM,"
	ssql=ssql+" ATUA_TX_OPERACAO,"
	ssql=ssql+" ATUA_CD_NR_USUARIO,"
	ssql=ssql+" ATUA_DT_ATUALIZACAO"	
	ssql=ssql+" FROM GRADE_MULTIPLICADOR "
	ssql=ssql+"WHERE CORT_CD_CORTE=" & Corte_Anterior & " AND MULT_TX_TIPO_MULTIPLICADOR<>'DESCENTRALIZADO')" 
	
	db_banco.execute(ssql)
	
	ssql="SELECT DISTINCT "
	ssql=ssql+"APOIO_LOCAL_MULT.USMA_CD_USUARIO, APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, "
	ssql=ssql+"APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR,APOIO_LOCAL_MULT.APLO_NR_SITUACAO, APOIO_LOCAL_MULT.APLO_NR_RELACAO_EMPREGO, "
	ssql=ssql+"APOIO_LOCAL_MULT.ATUA_TX_OPERACAO,APOIO_LOCAL_MULT.ATUA_CD_NR_USUARIO,APOIO_LOCAL_MULT.ATUA_DT_ATUALIZACAO, "
	ssql=ssql+"USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO "
	ssql=ssql+"FROM APOIO_LOCAL_MULT "
	ssql=ssql+"INNER JOIN USUARIO_MAPEAMENTO ON "
	ssql=ssql+"APOIO_LOCAL_MULT.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=2 AND APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1 "
	
	set fonte = db_banco.execute(ssql)
	
	set maximo = db_banco.execute("SELECT MAX(MULT_NR_CD_ID_MULT) AS CODIGO FROM GRADE_MULTIPLICADOR WHERE CORT_CD_CORTE=" & Corte_Atual)

	if isnull(maximo("CODIGO")) then
		Codigo_mult = 1
	else				
		Codigo_mult = maximo("CODIGO") + 1
	end if
	
	do until fonte.eof=true
	
		ssql=""
		ssql="INSERT INTO GRADE_MULTIPLICADOR("
		ssql=ssql+"CORT_CD_CORTE,"
		ssql=ssql+" MULT_NR_CD_ID_MULT,"
		ssql=ssql+" MULT_NR_TIPO_MULTIPLICADOR,"
		ssql=ssql+" MULT_TX_TIPO_MULTIPLICADOR,"
		ssql=ssql+" MULT_NR_CD_CHAVE,"
		ssql=ssql+" MULT_TX_NOME_MULTIPLICADOR,"
		ssql=ssql+" ORME_CD_ORG_MENOR,"
		ssql=ssql+" APLO_NR_RELACAO_EMPREGO,"
		ssql=ssql+" MULT_TX_RESTRICAO_VIAGEM,"
		ssql=ssql+" ATUA_TX_OPERACAO,"
		ssql=ssql+" ATUA_CD_NR_USUARIO,"
		ssql=ssql+" ATUA_DT_ATUALIZACAO)"	
		ssql=ssql+" VALUES ( "
		ssql=ssql+"" & Corte_Atual & ", "
		ssql=ssql+"" & Codigo_mult & ","
		ssql=ssql+" 1,"
		ssql=ssql+" 'DESCENTRALIZADO',"
		ssql=ssql+" '" & fonte("USMA_CD_USUARIO") & "',"
		ssql=ssql+" '" & fonte("USMA_TX_NOME_USUARIO") & "',"
		ssql=ssql+" '" & RIGHT("0000000" & fonte("ORME_CD_ORG_MENOR"),15) & "',"
		ssql=ssql+" '" & fonte("APLO_NR_RELACAO_EMPREGO") & "',"
		ssql=ssql+" NULL,"
		ssql=ssql+" 'I',"
		ssql=ssql+" '" & fonte("ATUA_CD_NR_USUARIO") & "',"
		ssql=ssql+"GETDATE())"
		
		db_banco.Execute(ssql)
		
		Codigo_mult = Codigo_mult + 1
		
		fonte.movenext
	
	loop
	
	'======== COPIA CURSOS REFERENTES À MULTIPLICADORES NO CORTE ATUAL ===============	
	
	set fonte_mult = db_banco.execute("SELECT DISTINCT MULT_NR_CD_ID_MULT, MULT_NR_CD_CHAVE FROM GRADE_MULTIPLICADOR WHERE CORT_CD_CORTE=" & Corte_Atual)
	
	do until fonte_mult.eof=true
	
	set fonte_curso = db_banco.execute("SELECT DISTINCT CURS_CD_CURSO FROM APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO='" & fonte_mult("MULT_NR_CD_CHAVE") & "'")

	if fonte_curso.eof=true then
		set fonte_curso = db_banco.execute("SELECT DISTINCT CURS_CD_CURSO FROM GRADE_MULTIPLICADOR_CURSO WHERE MULT_NR_CD_ID_MULT='" & fonte_mult("MULT_NR_CD_ID_MULT") & "' AND CORT_CD_CORTE=" & Corte_Anterior)	
	end if
	
		do until fonte_curso.eof=true
	
			ssql=""
			ssql="INSERT INTO GRADE_MULTIPLICADOR_CURSO("
			ssql=ssql+"CORT_CD_CORTE,"
			ssql=ssql+" MULT_NR_CD_ID_MULT,"
			ssql=ssql+" CURS_CD_CURSO,"
			ssql=ssql+" ATUA_TX_OPERACAO,"
			ssql=ssql+" ATUA_CD_NR_USUARIO,"
			ssql=ssql+" ATUA_DT_ATUALIZACAO)"	
			ssql=ssql+"VALUES( "
			ssql=ssql+"" & Corte_Atual & ", "
			ssql=ssql+"" & fonte_mult("MULT_NR_CD_ID_MULT") & " ,"
			ssql=ssql+"'" & fonte_curso("CURS_CD_CURSO") & "',"
			ssql=ssql+" 'I',"
			ssql=ssql+" 'XD47',"
			ssql=ssql+"GETDATE())"
	
			db_banco.Execute(ssql)
	
			fonte_curso.movenext
	
		loop
		
		fonte_mult.movenext
		
	loop

	'======== COPIA ÓRGÃOS REFERENTES À MULTIPLICADORES NO CORTE ATUAL ===============	
	
	set fonte_mult = db_banco.execute("SELECT DISTINCT MULT_NR_CD_ID_MULT, MULT_NR_CD_CHAVE FROM GRADE_MULTIPLICADOR WHERE CORT_CD_CORTE=" & Corte_Atual)
	
	do until fonte_mult.eof=true
	
	set fonte_orgao = db_banco.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO='" & fonte_mult("MULT_NR_CD_CHAVE") & "' AND APLO_NR_ATRIBUICAO=2")
	
	if fonte_orgao.eof=true then
		set fonte_orgao = db_banco.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM GRADE_MULTIPLICADOR_ORGAO_MENOR WHERE MULT_NR_CD_ID_MULT='" & fonte_mult("MULT_NR_CD_ID_MULT") & "' AND CORT_CD_CORTE=" & Corte_Anterior)	
	end if
	
		do until fonte_orgao.eof=true
	
			ssql=""
			ssql="INSERT INTO GRADE_MULTIPLICADOR_ORGAO_MENOR("
			ssql=ssql+"CORT_CD_CORTE,"
			ssql=ssql+" MULT_NR_CD_ID_MULT,"
			ssql=ssql+" ORME_CD_ORG_MENOR,"
			ssql=ssql+" ATUA_TX_OPERACAO,"
			ssql=ssql+" ATUA_CD_NR_USUARIO,"
			ssql=ssql+" ATUA_DT_ATUALIZACAO)"	
			ssql=ssql+"VALUES( "
			ssql=ssql+"" & Corte_Atual & ", "
			ssql=ssql+"" & fonte_mult("MULT_NR_CD_ID_MULT") & " ,"
			ssql=ssql+"'" &  RIGHT("0000000" &  fonte_orgao("ORME_CD_ORG_MENOR"),15) & "',"
			ssql=ssql+" 'I',"
			ssql=ssql+" 'XD47',"
			ssql=ssql+"GETDATE())"
			
			db_banco.Execute(ssql)
	
			fonte_orgao.movenext
	
		loop
		
		fonte_mult.movenext
		
	loop

	
	'======== CARREGA DEMANDA E UNIDADES PARA O NOVO CORTE ===============	
	
    str_SQL = ""
    str_SQL = str_SQL & " SELECT DISTINCT"
    str_SQL = str_SQL & " dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT,"
    str_SQL = str_SQL & " dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.CURSO_FUNCAO.CURS_CD_CURSO,"
    str_SQL = str_SQL & " dbo.CURSO.CURS_NUM_CARGA_CURSO, dbo.CURSO.CURS_TX_METODO_CURSO, COUNT(DISTINCT dbo.FUNCAO_USUARIO.USMA_CD_USUARIO)"
    str_SQL = str_SQL & " AS Total"
    str_SQL = str_SQL & " FROM         dbo.ORGAO_MAIOR INNER JOIN"
    str_SQL = str_SQL & " dbo.ORGAO_MENOR ON dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT = dbo.ORGAO_MENOR.ORLO_CD_ORG_LOT INNER JOIN"
    str_SQL = str_SQL & " dbo.USUARIO_MAPEAMENTO ON dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR = dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR INNER JOIN"
    str_SQL = str_SQL & " dbo.FUNCAO_USUARIO ON dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO = dbo.FUNCAO_USUARIO.USMA_CD_USUARIO INNER JOIN"
    str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO ON dbo.FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
    str_SQL = str_SQL & " dbo.CURSO_FUNCAO ON dbo.FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
    str_SQL = str_SQL & " dbo.CURSO ON dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.CURSO.CURS_CD_CURSO INNER JOIN"
    str_SQL = str_SQL & " dbo.ABRANGENCIA_CURSO ON dbo.CURSO.ONDA_CD_ONDA = dbo.ABRANGENCIA_CURSO.ONDA_CD_ONDA INNER JOIN"
    str_SQL = str_SQL & " dbo.USUARIO_APROVADO ON dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.USUARIO_APROVADO.CURS_CD_CURSO"
    str_SQL = str_SQL & " WHERE dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT NOT IN (886,241,240)"
    str_SQL = str_SQL & " and dbo.CURSO.CURS_TX_STATUS_CURSO = '1'"
    str_SQL = str_SQL & " GROUP BY dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT, dbo.CURSO_FUNCAO.CURS_CD_CURSO, dbo.ORGAO_MAIOR.ORLO_CD_STATUS,"
    str_SQL = str_SQL & " dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.ORGAO_MENOR.ORME_CD_STATUS, dbo.FUNCAO_USUARIO.FUUS_IN_PRIORITARIO,"
    str_SQL = str_SQL & " dbo.CURSO.CURS_NUM_CARGA_CURSO, dbo.CURSO.CURS_TX_METODO_CURSO, dbo.ABRANGENCIA_CURSO.ONDA_CD_ONDA,"
    str_SQL = str_SQL & " dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO, dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT,"
    str_SQL = str_SQL & " dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR"
    str_SQL = str_SQL & " HAVING (dbo.ORGAO_MAIOR.ORLO_CD_STATUS = 'A') AND (dbo.ORGAO_MENOR.ORME_CD_STATUS = 'A') AND"
    str_SQL = str_SQL & " (dbo.FUNCAO_USUARIO.FUUS_IN_PRIORITARIO = '1') AND (dbo.ABRANGENCIA_CURSO.ONDA_CD_ONDA = 6 OR"
    str_SQL = str_SQL & " dbo.ABRANGENCIA_CURSO.ONDA_CD_ONDA = 9) AND (NOT (dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO = 'LM')) AND"
    str_SQL = str_SQL & " (NOT (dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO = 'AP'))"
    str_SQL = str_SQL & " ORDER BY dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.CURSO_FUNCAO.CURS_CD_CURSO"

    Set rst_Demanda = db_banco.Execute(str_SQL)
    
	int_RecodCount_Demanda = rst_Demanda.RecordCount
    
	Do While Not rst_Demanda.EOF
                                        
        str_SQL = ""
        str_SQL = str_SQL & " Insert into GRADE_DEMANDA_ORIGINAL_SEM("
        str_SQL = str_SQL & " CORT_CD_CORTE"
        str_SQL = str_SQL & " , ORLO_CD_ORG_LOT"
        str_SQL = str_SQL & " , ORME_CD_ORG_MENOR"
        str_SQL = str_SQL & " , CURS_CD_CURSO"
        str_SQL = str_SQL & " , DEMA_NR_TOTAL"
        str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
        str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
        str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
        str_SQL = str_SQL & " )Values("
        str_SQL = str_SQL & Corte_Atual
        str_SQL = str_SQL & "," & Trim(rst_Demanda("ORLO_CD_ORG_LOT"))
        str_SQL = str_SQL & ",'" & RIGHT("0000000"+ rst_Demanda("ORME_CD_ORG_MENOR"),15)
        str_SQL = str_SQL & "','" & Trim(rst_Demanda("CURS_CD_CURSO")) & "'"
        str_SQL = str_SQL & "," & Trim(rst_Demanda("Total"))
                        
        str_SQL = str_SQL & ",'C' ,'XK45' ,GETDATE())"
        
        Set rdsNovo = db_banco.Execute(str_SQL)
        
        int_Tot_Carga_Tot = int_Tot_Carga_Tot + 1
    
        rst_Demanda.MoveNext
        
    Loop

    str_SQL = ""
    str_SQL = str_SQL & " Select DISTINCT "
    str_SQL = str_SQL & " dbo.GRADE_DEMANDA_ORIGINAL_SEM.CORT_CD_CORTE"
    str_SQL = str_SQL & " , dbo.GRADE_DEMANDA_ORIGINAL_SEM.ORLO_CD_ORG_LOT"
    str_SQL = str_SQL & " , dbo.GRADE_ORGAO_MAIOR.ORLO_SG_ORG_LOT "
    str_SQL = str_SQL & " , dbo.GRADE_ORGAO_MAIOR.ORLO_CD_GABINETE"
    str_SQL = str_SQL & " FROM  dbo.GRADE_DEMANDA_ORIGINAL_SEM INNER JOIN"
    str_SQL = str_SQL & " dbo.GRADE_ORGAO_MAIOR ON"
    str_SQL = str_SQL & " dbo.GRADE_DEMANDA_ORIGINAL_SEM.ORLO_CD_ORG_LOT = dbo.GRADE_ORGAO_MAIOR.ORLO_CD_ORG_LOT"
    str_SQL = str_SQL & " WHERE (dbo.GRADE_DEMANDA_ORIGINAL_SEM.CORT_CD_CORTE = " & Corte_Atual & ")"
    str_SQL = str_SQL & " ORDER BY dbo.GRADE_ORGAO_MAIOR.ORLO_SG_ORG_LOT"
    
    Set rst_Orgao_Maior = db_banco.Execute(str_SQL)
        
    Do While Not rst_Orgao_Maior.EOF
            
        str_SQL = ""
        str_SQL = str_SQL & " SELECT"
        str_SQL = str_SQL & " MAX(UNID_CD_UNIDADE) As CODIGO"
        str_SQL = str_SQL & " FROM GRADE_UNIDADE"
        
		Set rstMax = db_banco.Execute(str_SQL)
        
		If Not IsNull(rstMax("codigo")) Then
          int_Prox_Unidade = rstMax("CODIGO") + 1
        Else
           int_Prox_Unidade = 1
        End If
        rstMax.Close
        Set rstMax = Nothing
                                                                                        
        str_SQL = ""
        str_SQL = str_SQL & " Select "
        str_SQL = str_SQL & " CTRO_CD_CENTRO_TREINAMENTO"
        str_SQL = str_SQL & " from GRADE_UNIDADE"
        str_SQL = str_SQL & " where UNID_TX_DESC_UNIDADE='" & Trim(rst_Orgao_Maior("ORLO_SG_ORG_LOT")) & "'"
        str_SQL = str_SQL & " AND CORT_CD_CORTE = " & Corte_Anterior
		        
		Set rst_CdCT = db_banco.Execute(str_SQL)
        
		If Not rst_CdCT.EOF Then
            int_Cd_CT = rst_CdCT("CTRO_CD_CENTRO_TREINAMENTO")
        Else
            int_Cd_CT = 0
        End If
    
	    rst_CdCT.Close
        Set rst_CdCT = Nothing
                                                                                    
        str_SQL = ""
        str_SQL = str_SQL & " Insert into GRADE_UNIDADE("
        str_SQL = str_SQL & " UNID_CD_UNIDADE"
        str_SQL = str_SQL & " , ORLO_CD_ORG_LOT_DIR"
        str_SQL = str_SQL & " , UNID_TX_DESC_UNIDADE"
        str_SQL = str_SQL & " , CTRO_CD_CENTRO_TREINAMENTO"
        str_SQL = str_SQL & " , CORT_CD_CORTE"
        str_SQL = str_SQL & " , ORLO_CD_ORG_LOT"
        str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
        str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
        str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
        str_SQL = str_SQL & " )Values("
        str_SQL = str_SQL & int_Prox_Unidade
        str_SQL = str_SQL & "," & Trim(rst_Orgao_Maior("ORLO_CD_GABINETE"))
        str_SQL = str_SQL & ",'" & Trim(rst_Orgao_Maior("ORLO_SG_ORG_LOT")) & "'"
    
	    If int_Cd_CT <> 0 Then
            str_SQL = str_SQL & "," & int_Cd_CT
        Else
            str_SQL = str_SQL & ",Null"
        End If
    
	    str_SQL = str_SQL & "," & Corte_Atual
        str_SQL = str_SQL & "," & Trim(rst_Orgao_Maior("ORLO_CD_ORG_LOT"))
        str_SQL = str_SQL & ",'C' ,'XK45' ,GETDATE())"
        
        Set rdsNovo = db_banco.Execute(str_SQL)
    
        set rs_unidade = db_banco.execute("SELECT DISTINCT ORLO_CD_ORG_LOT FROM GRADE_UNIDADE WHERE CORT_CD_CORTE=" & Corte_Atual)
		
		do until rs_unidade.eof = true		
			
			str_SQL = ""
			str_SQL = str_SQL & " Select DISTINCT"
			str_SQL = str_SQL & " ORME_CD_ORG_MENOR"
			str_SQL = str_SQL & " FROM GRADE_DEMANDA_ORIGINAL_SEM"
			str_SQL = str_SQL & " WHERE CORT_CD_CORTE = " & Corte_Atual
			str_SQL = str_SQL & " AND ORLO_CD_ORG_LOT = " & Trim(rs_unidade("ORLO_CD_ORG_LOT"))
			
			Set rst_Orgao_Menor = db_banco.Execute(str_SQL)
			
				If Not rst_Orgao_Menor.EOF Then
			
					Do While Not rst_Orgao_Menor.EOF
					
					str_SQL = ""
					str_SQL = str_SQL & " Insert into GRADE_UNIDADE_ORGAO_MENOR("
					str_SQL = str_SQL & " CORT_CD_CORTE"
					str_SQL = str_SQL & " , UNID_CD_UNIDADE"
					str_SQL = str_SQL & " , ORME_CD_ORG_MENOR"
					str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
					str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
					str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
					str_SQL = str_SQL & " )Values("
					str_SQL = str_SQL & Corte_Atual & ", "
					str_SQL = str_SQL & int_Prox_Unidade
					str_SQL = str_SQL & ",'" & RIGHT("0000000"+ rst_Orgao_Menor("ORME_CD_ORG_MENOR"),15)
					str_SQL = str_SQL & "','C' ,'XK45' ,GETDATE())"
					
					Set rdsNovo = db_banco.Execute(str_SQL)
					
					rst_Orgao_Menor.MoveNext
				
					Loop
				
				End If
			
				rs_unidade.movenext
			
			loop
    	
		rst_Orgao_Maior.MoveNext
	
	Loop
	
	'============================== FIM DE CÓPIA DE CORTE ===========================================
	
'************************************** ALTERAÇÃO DE CORTE ************************************************	
elseif strAcao = "A" then			
		
	strCdCorte 	= trim(Request("txtCdCorte"))		
		
	strSQLAltCorte = ""
	strSQLAltCorte = strSQLAltCorte & "UPDATE GRADE_CORTE "
	strSQLAltCorte = strSQLAltCorte & "SET CORT_TX_DESC_CORTE = '" & strNomeCorte & "',"		
	strSQLAltCorte = strSQLAltCorte & "CORT_DT_DATA_CORTE = '" & strDtCorte & "',"		
	strSQLAltCorte = strSQLAltCorte & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltCorte = strSQLAltCorte & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltCorte = strSQLAltCorte & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltCorte = strSQLAltCorte & "WHERE CORT_CD_CORTE = " & strCdCorte 
	'Response.write strSQLAltCorte
	'Response.end	
	
	on error resume next
		db_banco.Execute(strSQLAltCorte)
					
	if err.number = 0 then		
		strMSG = "Corte foi alterado com sucesso."
	else
		strMSG = "Houve um erro na alteração do corte (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE CORTE ************************************************	
elseif strAcao = "E" then 
	
	strCdCorte 	= trim(Request("selCorte"))	
				
	strSQLDelCorte = ""
	strSQLDelCorte = strSQLDelCorte & "DELETE FROM GRADE_CORTE "	
	strSQLDelCorte = strSQLDelCorte & "WHERE CORT_CD_CORTE = " & strCdCorte 
	
	on error resume next

		db_banco.Execute(strSQLDelCorte)
		db_banco.execute("DELETE FROM GRADE_CENTRO_TREINAMENTO WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_SALA WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_FERIADO_SALA WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_CURSO WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_CURSO_PRE_REQUISITO WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_ORGAO_MAIOR WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_ORGAO_MENOR WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_MULTIPLICADOR WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_MULTIPLICADOR_CURSO WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_MULTIPLICADOR_ORGAO_MENOR WHERE CORT_CD_CORTE=" & strCdCorte)		
		db_banco.execute("DELETE FROM GRADE_DEMANDA_ORIGINAL_SEM WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_UNIDADE WHERE CORT_CD_CORTE=" & strCdCorte)							
		db_banco.execute("DELETE FROM GRADE_UNIDADE_ORGAO_MENOR WHERE CORT_CD_CORTE=" & strCdCorte)
		db_banco.execute("DELETE FROM GRADE_CURSO_UNIDADE WHERE CORT_CD_CORTE=" & strCdCorte)									
	
	if err.number = 0 then		
		strMSG = "Corte foi excluído com sucesso."
	else
		strMSG = "Houve um erro na exclusão do corte (" & err.description & ")"
	end if	

end if

public function MontaDataHora(strData,intDataTime)

	'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
	'*** intDataTime = 1 (DATA E HORA)
	'*** intDataTime = 2 (DATA)
	'*** intDataTime = 3 (HORA)
	'*** intDataTime = 4 (FORMATO DE BANCO)

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
	end if
end function
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
			  <table border="0" width="849" height="139">
					  <tr>
						
				  <td width="205" height="29"></td>
						
				  <td width="93" height="29" valign="middle" align="left"></td>
						
				  <td height="29" valign="middle" align="left" colspan="2"> 				 
				  	<b><font face="Verdana" color="#330099" size="2"><%=strMSG%></font></b> 
				  </td>
						
				  </tr>
			  
				  <tr>
					<td height="21"></td>
					<td height="21" valign="middle" align="left"></td>
					<td height="21" valign="middle" align="left">&nbsp;</td>
					<td height="21" valign="middle" align="left">&nbsp;</td>
			    </tr>
				<tr>
						
				  <td width="205" height="35"></td>
						
				  <td width="93" height="35" valign="middle" align="left"></td>
						
				  <td height="35" valign="middle" align="left" width="29"> 
					<a href="../../indexA_grade.asp?selCorte=<%=strSala%>"> 
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a>
				 </td>	
				 <td height="35" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font>
				 </td>
				</tr>
					
				<tr>						
				  <td width="205" height="29"></td>						
				  <td width="93" height="29" valign="middle" align="left"></td>						
				  <td height="29" valign="middle" align="left" width="29"> 				 
				   	<a href="sel_corte.asp">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Cortes</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
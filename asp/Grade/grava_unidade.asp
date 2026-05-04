<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if trim(Request("parAcao"))	<> "" then
	strAcao = trim(Request("parAcao"))	
elseif trim(Request("pAcao"))	<> "" then
	strAcao = trim(Request("pAcao"))	
end if

if trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))	
end if		

intCdUnidade = trim(Request("selUnidade"))	

intCdDiretoria = trim(Request("selDiretoria"))
intCdCT = trim(Request("selCT"))
strNomeUnidade = Ucase(trim(Request("txtNomeUnidade")))	
strOrgSel = trim(Request("txtOrgSel"))

'response.write "intCdDiretoria = " & intCdDiretoria & "<br>"	
'response.write "intCdCT = " & intCdCT & "<br>"	
'response.write "strNomeUnidade = " & strNomeUnidade & "<br>"	
'response.write "strOrgSel = " & strOrgSel & "<br><br>"
'response.end		

if strAcao = "I" then
	strNomeAcao = "Inclusão de Unidade"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Unidade"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Unidade"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE UNIDADE ************************************************
if strAcao = "I" then	
	
	strVerificaUnidade = ""
	strVerificaUnidade = strVerificaUnidade & "SELECT UNID_TX_DESC_UNIDADE "
	strVerificaUnidade = strVerificaUnidade & "FROM GRADE_UNIDADE "
	strVerificaUnidade = strVerificaUnidade & "WHERE UNID_TX_DESC_UNIDADE = '" & strNomeUnidade & "'"	
	strVerificaUnidade = strVerificaUnidade & " AND CORT_CD_CORTE = " & Session("Corte")
	'Response.write strVerificaUnidade
	'Response.end
		
	Set rdsVerificaUnidade = db_banco.Execute(strVerificaUnidade)			
	
	if not rdsVerificaUnidade.EOF then
		strMSG = "Já existe unidade cadastrada com o nome - " & rdsVerificaUnidade("UNID_TX_DESC_UNIDADE") & " para o Corte selecionado."
	else			
	
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(UNID_CD_UNIDADE) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_UNIDADE "		
		strVerificaCod = strVerificaCod & "WHERE CORT_CD_CORTE = " & Session("Corte")	
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCdUnidade = 1
			else
				intCdUnidade = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCdUnidade = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
				
		strSQLIncUnidade = ""
		strSQLIncUnidade = strSQLIncUnidade & "INSERT INTO GRADE_UNIDADE (CORT_CD_CORTE, UNID_CD_UNIDADE, CTRO_CD_CENTRO_TREINAMENTO, "
		strSQLIncUnidade = strSQLIncUnidade & "ORLO_CD_ORG_LOT_DIR, UNID_TX_DESC_UNIDADE, "
		strSQLIncUnidade = strSQLIncUnidade & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncUnidade = strSQLIncUnidade & "VALUES(" & Session("Corte") & "," & intCdUnidade & "," & intCdCT & "," & intCdDiretoria & ",'" & strNomeUnidade 
		strSQLIncUnidade = strSQLIncUnidade & "','I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLIncUnidade & "<br><br>"
		'Response.end	
				
		on error resume next
			db_banco.Execute(strSQLIncUnidade)	
			
		'*** CADASTRA AS NOVOS REGISTROS ***			
		vetCDsOrgao = split(strOrgSel,",")
					
		r = 0			
		for r = lbound(vetCDsOrgao) to Ubound(vetCDsOrgao)				
						
			if vetCDsOrgao(r) <> "" and vetCDsOrgao(r) <> "0" then				
			
				intCdOrgaoResult = cstr(vetCDsOrgao(r))
						
				strSQLIncUnidadeOrgaoMenor = ""		
				strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "INSERT INTO GRADE_UNIDADE_ORGAO_MENOR "
				strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "(CORT_CD_CORTE, UNID_CD_UNIDADE, ORME_CD_ORG_MENOR, "
				strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
				strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "VALUES(" & Session("Corte") & "," & intCdUnidade & ",'" & intCdOrgaoResult
				strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "','I','" & Session("CdUsuario") & "',GETDATE())"					
				'response.write strSQLIncUnidadeOrgaoMenor & "<br><br>"
				'Response.end					
				db_banco.Execute(strSQLIncUnidadeOrgaoMenor)
			end if
		next		
			
		if err.number = 0 then		
			strMSG = "Unidade foi incluída com sucesso."
		else
			'if err.number = "2147217873" then		
				strMSG = "Houve um erro na inclusão da unidade (" & err.description & " - " & err.number & ")"
			'end if
		end if	
	end if
	
	rdsVerificaUnidade.close
	set rdsVerificaUnidade = nothing
	
'************************************** ALTERAÇÃO DE UNIDADE ************************************************	
elseif strAcao = "A" then			
						
	'*** LIMPA OS REGISTROS DA TABELA GRADE_UNIDADE_ORGAO_MENOR **
	strSQLDelUnidadeOrgaoMenor = ""
	strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & "DELETE FROM GRADE_UNIDADE_ORGAO_MENOR "	
	strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & "WHERE UNID_CD_UNIDADE = " & intCdUnidade 
	strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & " AND CORT_CD_CORTE = " & Session("Corte")
	'response.write strSQLDelUnidadeOrgaoMenor
	'Response.end		
	
	db_banco.Execute(strSQLDelUnidadeOrgaoMenor)	
				
	'*** ALTERA AS NOVOS REGISTROS ***			
	vetCDsOrgao = split(strOrgSel,",")
				
	r = 0			
	for r = lbound(vetCDsOrgao) to Ubound(vetCDsOrgao)				
		
		'Response.write vetCDsOrgao(r) & "<br><br>"		
		
		if vetCDsOrgao(r) <> "" then				
			intCdOrgaoResult = vetCDsOrgao(r)
					
			strSQLIncUnidadeOrgaoMenor = ""		
			strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "INSERT INTO GRADE_UNIDADE_ORGAO_MENOR "
			strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "(CORT_CD_CORTE, UNID_CD_UNIDADE, ORME_CD_ORG_MENOR, "
			strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
			strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "VALUES(" & Session("Corte") & "," & intCdUnidade & ",'" & intCdOrgaoResult 
			strSQLIncUnidadeOrgaoMenor = strSQLIncUnidadeOrgaoMenor & "','I','" & Session("CdUsuario") & "',GETDATE())"					
			'response.write strSQLIncUnidadeOrgaoMenor & "<br><br>"
			'Response.end					
			db_banco.Execute(strSQLIncUnidadeOrgaoMenor)
		end if
	next						
		
	strSQLAltFeriado = ""
	strSQLAltFeriado = strSQLAltFeriado & "UPDATE GRADE_UNIDADE "
	strSQLAltFeriado = strSQLAltFeriado & "SET CTRO_CD_CENTRO_TREINAMENTO = " & intCdCT & ","	
	strSQLAltFeriado = strSQLAltFeriado & "ORLO_CD_ORG_LOT_DIR = " & intCdDiretoria & ","	
	strSQLAltFeriado = strSQLAltFeriado & "UNID_TX_DESC_UNIDADE = '" & strNomeUnidade & "',"		
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltFeriado = strSQLAltFeriado & "WHERE UNID_CD_UNIDADE = " & intCdUnidade 	
	strSQLAltFeriado = strSQLAltFeriado & " AND CORT_CD_CORTE = " & Session("Corte")
	'Response.write strSQLAltFeriado
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltFeriado)
					
	if err.number = 0 then		
		strMSG = "Unidade foi alterada com sucesso."
	else
		strMSG = "Houve um erro na alteração da unidade (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE UNIDADE ************************************************	
elseif strAcao = "E" then 
				
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmUnidade = ""	
	
	'*** VERIFICA EM ORGÃO MENOR ****
	strSQLVeriUnidadeOrgMenor = ""			
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & "SELECT UNIDADE.UNID_CD_UNIDADE, UNIDADE.UNID_TX_DESC_UNIDADE "
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & "FROM GRADE_UNIDADE UNIDADE, GRADE_UNIDADE_ORGAO_MENOR ORG_MENOR "
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & "WHERE UNIDADE.UNID_CD_UNIDADE = ORG_MENOR.UNID_CD_UNIDADE "
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & "AND UNIDADE.UNID_CD_UNIDADE = " & intCdUnidade 
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte") 
	strSQLVeriUnidadeOrgMenor = strSQLVeriUnidadeOrgMenor & " AND ORG_MENOR.CORT_CD_CORTE = " & Session("Corte") 		
	'Response.write strSQLVeriUnidadeOrgMenor & "<br>"
	'Response.end		
			
	set rsVeriUnidadeOrgMenor = db_banco.Execute(strSQLVeriUnidadeOrgMenor)		
			
	if not rsVeriUnidadeOrgMenor.eof then	
		blnPodeDeletar = false
		strAssociacao = "Orgão Menor"
		strNmUnidade = rsVeriUnidadeOrgMenor("UNID_TX_DESC_UNIDADE")
	end if
	
	'*** VERIFICA EM CURSO UNIDADE ****			
	strSQLVeriCursoUnid = ""			
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & "SELECT UNIDADE.UNID_CD_UNIDADE, UNIDADE.UNID_TX_DESC_UNIDADE "
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & "FROM GRADE_UNIDADE UNIDADE, GRADE_CURSO_UNIDADE CURSO_UNID "
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & "WHERE UNIDADE.UNID_CD_UNIDADE = CURSO_UNID.UNID_CD_UNIDADE "
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & "AND UNIDADE.UNID_CD_UNIDADE = " & intCdUnidade 
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte") 
	strSQLVeriCursoUnid = strSQLVeriCursoUnid & " AND CURSO_UNID.CORT_CD_CORTE = " & Session("Corte") 		
	'Response.write strSQLVeriCursoUnid & "<br>"
	'Response.end		
			
	set rsVeriCursoUnid = db_banco.Execute(strSQLVeriCursoUnid)		
			
	if not rsVeriCursoUnid.eof then	
		blnPodeDeletar = false
		if strAssociacao <> "" then
			strAssociacao = "Curso"
		else
			strAssociacao = strAssociacao & ", Curso "
		end if
		strNmUnidade = rsVeriCursoUnid("UNID_TX_DESC_UNIDADE")
	end if							
				
	'*** VERIFICA EM CURSO TURMA ****			
	strSQLVeriCursoTurma = ""			
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & "SELECT UNIDADE.UNID_CD_UNIDADE, UNIDADE.UNID_TX_DESC_UNIDADE "
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & "FROM GRADE_UNIDADE UNIDADE, GRADE_TURMA_UNIDADE TURMA_UNID "
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & "WHERE UNIDADE.UNID_CD_UNIDADE = TURMA_UNID.UNID_CD_UNIDADE "
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & "AND UNIDADE.UNID_CD_UNIDADE = " & intCdUnidade 
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte") 
	strSQLVeriCursoTurma = strSQLVeriCursoTurma & " AND TURMA_UNID.CORT_CD_CORTE = " & Session("Corte") 		
	'Response.write strSQLVeriCursoTurma & "<br>"
	'Response.end		
			
	set rsVeriCursoTurma = db_banco.Execute(strSQLVeriCursoTurma)		
			
	if not rsVeriCursoTurma.eof then	
		blnPodeDeletar = false
		if strAssociacao <> "" then
			strAssociacao = "Turma"
		else
			strAssociacao = strAssociacao & ", Turma "
		end if
		strNmUnidade = rsVeriCursoTurma("UNID_TX_DESC_UNIDADE")
	end if										
				
	'Response.write blnPodeDeletar & " - MSG - " & strAssociacao
	'Response.end					
				
	if blnPodeDeletar then
				
		'*** LIMPA OS REGISTROS DA TABELA PARA O FERIADO SELECIONADO **
		strSQLDelUnidadeOrgaoMenor = ""
		strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & "DELETE FROM GRADE_UNIDADE_ORGAO_MENOR "	
		strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & "WHERE UNID_CD_UNIDADE = " & intCdUnidade 
		strSQLDelUnidadeOrgaoMenor = strSQLDelUnidadeOrgaoMenor & " AND CORT_CD_CORTE = " & Session("Corte")
		'response.write strSQLDelUnidadeOrgaoMenor
		'Response.end		
		
		db_banco.Execute(strSQLDelUnidadeOrgaoMenor)				
			
		strSQLDelFeriado = ""
		strSQLDelFeriado = strSQLDelFeriado & "DELETE FROM GRADE_UNIDADE "		
		strSQLDelFeriado = strSQLDelFeriado & "WHERE UNID_CD_UNIDADE = " & intCdUnidade
		strSQLDelFeriado = strSQLDelFeriado & " AND CORT_CD_CORTE = " & Session("Corte")
		'Response.write strSQLDelFeriado
		'Response.end
		
		on error resume next
			db_banco.Execute(strSQLDelFeriado)
		
		if err.number = 0 then		
			strMSG = "Unidade foi excluída com sucesso."
		else
			strMSG = "Houve um erro na exclusão da unidade (" & err.description & ")"
		end if	
	else
		strMSG = "A Unidade - " & strNmUnidade & " não pode ser excluída, pois existe associação da mesma em " & strAssociacao & "."
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
				   	<a href="sel_unidade.asp?selCorte=<%=Session("Corte")%>" >			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Unidade</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
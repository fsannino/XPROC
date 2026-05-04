<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

'strUnidDir		= trim(Request("txtUnidDir"))
strSala 		= trim(Request("hdSala"))
strAcao 		= trim(Request("parAcao"))
strCurso 		= trim(Request("selCurso"))
intCDTurma 		= trim(Request("txtCDTurma"))
strNomeTurma 	= Ucase(trim(Request("txtNomeTurma")))
strMultiplic 	= trim(Request("selMultiplic"))
strCorte	 	= session("Corte")
strMandante 	= Ucase(trim(Request("txtMandante")))
intCdsUnid_Sel  = trim(Request("txtUnidades_Selecionadas"))	
	
if trim(Request("txtDtIni")) <> "" then
	strDtIni 		= MontaDataHora(trim(Request("txtDtIni")),4)
else
	strDtIni = ""
end if

if trim(Request("txtDtFim")) <> "" then
	strDtFim 		= MontaDataHora(trim(Request("txtDtFim")),4)
else
	strDtFim = ""
end if

strHrIni 		= trim(Request("txtHrIni"))
strHrFim 		= trim(Request("txtHrFim"))
'strQtdePeriodo 	= trim(Request("txtQtdePeriodo"))
	
'Response.write "strAcao - " & strAcao & "<br>"
'Response.write "strTipo - " & strTipo & "<br><br>"
''Response.write "strUnidDir - " & strUnidDir & "<br>"
'Response.write "strCurso - " & strCurso & "<br>"
'Response.write "strSala - " & strSala & "<br>"
'Response.write "intCDTurma - " & intCDTurma & "<br>"
'Response.write "strNomeTurma - " & strNomeTurma & "<br>"
'Response.write "strMultiplic - " & strMultiplic & "<br>"
'Response.write "strCorte - " & strCorte & "<br>"
'Response.write "strMandante - " & strMandante & "<br>"	
'Response.write "strDtIni - " & strDtIni & "<br>"
'Response.write "strDtFim - " & strDtFim & "<br>"
'Response.write "strHrIni - " & strHrIni & "<br>"
'Response.write "strHrFim - " & strHrFim & "<br>"
'Response.write "strQtdePeriodo - " & strQtdePeriodo & "<br><br>"
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Turma"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Turma"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Turma"
end if
	
strMSG =  ""			
	
'************************************** ALTERAÇÃO DE TURMA ************************************************	
if strAcao = "A" then			
		
	strVerificaCurso = ""	
	strVerificaCurso = strVerificaCurso & "SELECT CURS_NUM_CARGA_CURSO "
	strVerificaCurso = strVerificaCurso & "FROM CURSO "
	strVerificaCurso = strVerificaCurso & "WHERE CURS_CD_CURSO = '" & strCurso & "'"
		
	Set rdsVerificaCurso = db_banco.Execute(strVerificaCurso)			
	
	if not rdsVerificaCurso.EOF then				
		intParcialPeriodo = cint(rdsVerificaCurso("CURS_NUM_CARGA_CURSO") / 4)	
		strQtdePeriodo = intParcialPeriodo
	else
		strQtdePeriodo = 0
	end if		
		
	rdsVerificaCurso.close
	set rdsVerificaCurso = nothing	
		
	strSQLAltTurma = ""
	strSQLAltTurma = strSQLAltTurma & "UPDATE GRADE_TURMA SET "
	strSQLAltTurma = strSQLAltTurma & "TURM_TX_MANDANTE = '" & strMandante & "',"		
	'strSQLAltTurma = strSQLAltTurma & "SALA_CD_SALA = " & strSala & "," 
	strSQLAltTurma = strSQLAltTurma & "CURS_CD_CURSO = '" & strCurso & "',"
	'strSQLAltTurma = strSQLAltTurma & "USMA_CD_USUARIO = '" & strMultiplic & "',"	
	strSQLAltTurma = strSQLAltTurma & "MULT_NR_CD_ID_MULT = '" & strMultiplic & "',"	
	strSQLAltTurma = strSQLAltTurma & "TURM_TX_DESC_TURMA = '" & strNomeTurma & "',"	
	strSQLAltTurma = strSQLAltTurma & "CORT_CD_CORTE = " & Session("Corte") & ","		
	strSQLAltTurma = strSQLAltTurma & "TURM_DT_INICIO = '" & strDtIni & "',"	
	strSQLAltTurma = strSQLAltTurma & "TURM_DT_TERMINO = '" & strDtFim & "',"	
	strSQLAltTurma = strSQLAltTurma & "TURM_HR_INICIO = '" & strHrIni & "',"		
	strSQLAltTurma = strSQLAltTurma & "TURM_HR_TERMINO = '" & strHrFim & "',"	
	strSQLAltTurma = strSQLAltTurma & "TURM_NUM_QTE_PERIODO = " & strQtdePeriodo & ","
	strSQLAltTurma = strSQLAltTurma & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltTurma = strSQLAltTurma & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltTurma = strSQLAltTurma & "ATUA_DT_ATUALIZACAO = GETDATE() "		
	strSQLAltTurma = strSQLAltTurma & "WHERE TURM_NR_CD_TURMA = " & intCdTurma 	
	strSQLAltTurma = strSQLAltTurma & " AND CORT_CD_CORTE = " & Session("Corte")
	'Response.write strSQLAltTurma
	'Response.end			
		
	'*** LIMPA OS REGISTROS DA TABELA GRADE_TURMA_UNIDADE **
	strSQLDelTurmaUnid = ""
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & "DELETE FROM GRADE_TURMA_UNIDADE "	
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & "WHERE TURM_NR_CD_TURMA = " & intCdTurma 
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & " AND CORT_CD_CORTE = " & Session("Corte")
	'response.write strSQLDelTurmaUnid
	'Response.end		
	
	db_banco.Execute(strSQLDelTurmaUnid)	
				
	'*** ALTERA AS NOVOS REGISTROS ***			
	vetCDsUnidades = split(intCdsUnid_Sel,",")
				
	r = 0			
	for r = lbound(vetCDsUnidades) to Ubound(vetCDsUnidades)				
		
		'Response.write vetCDsUnidades(r) & "<br><br>"		
		
		if vetCDsUnidades(r) <> "" then				
			intCdUnidadeResult = cint(vetCDsUnidades(r))
					
			strSQLIncTurmaUnid = ""		
			strSQLIncTurmaUnid = strSQLIncTurmaUnid & "INSERT INTO GRADE_TURMA_UNIDADE "
			strSQLIncTurmaUnid = strSQLIncTurmaUnid & "(CORT_CD_CORTE, TURM_NR_CD_TURMA, UNID_CD_UNIDADE, "
			strSQLIncTurmaUnid = strSQLIncTurmaUnid & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
			strSQLIncTurmaUnid = strSQLIncTurmaUnid & "VALUES(" & Session("Corte") & "," & intCdTurma & "," & intCdUnidadeResult & "," 
			strSQLIncTurmaUnid = strSQLIncTurmaUnid & "'I','" & Session("CdUsuario") & "',GETDATE())"					
			'response.write strSQLIncTurmaUnid & "<br><br>"
			'Response.end					
			db_banco.Execute(strSQLIncTurmaUnid)
		end if
	next						
	
	on error resume next
		db_banco.Execute(strSQLAltTurma)
					
	if err.number = 0 then		
		strMSG = "Turma foi alterada com sucesso."
	else
		strMSG = "Houve um erro na alteração da turma (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE TURMA ************************************************	
elseif strAcao = "E" then 
				
	'*** LIMPA OS REGISTROS DA TABELA GRADE_TURMA_UNIDADE **
	strSQLDelTurmaUnid = ""
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & "DELETE FROM GRADE_TURMA_UNIDADE "	
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & "WHERE TURM_NR_CD_TURMA = " & intCdTurma 
	strSQLDelTurmaUnid = strSQLDelTurmaUnid & " AND CORT_CD_CORTE = " & Session("Corte")
	'response.write strSQLDelTurmaUnid
	'Response.end		
	
	db_banco.Execute(strSQLDelTurmaUnid)				
				
	strSQLDelTurma = ""
	strSQLDelTurma = strSQLDelTurma & "DELETE FROM GRADE_TURMA "	
	strSQLDelTurma = strSQLDelTurma & "WHERE TURM_NR_CD_TURMA = " & intCdTurma  
	strSQLDelTurma = strSQLDelTurma & " AND CORT_CD_CORTE = " & Session("Corte")				
	'Response.write strSQLDelTurma
	'Response.end
	
	on error resume next
		db_banco.Execute(strSQLDelTurma)
	
	if err.number = 0 then		
		strMSG = "Turma foi excluída com sucesso."
	else
		strMSG = "Houve um erro na exclusão da turma (" & err.description & ")"
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
					<a href="Inclui_altera_turma.asp?selSala=<%=strSala%>"> 
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="35" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
				</tr>
					
				<tr>						
				  <td width="205" height="29"></td>						
				  <td width="93" height="29" valign="middle" align="left"></td>						
				  <td height="29" valign="middle" align="left" width="29"> 				 
				   	<a href="inclui_altera_turma.asp?selSala=<%=strSala%>&selCorte=<%=Session("Corte")%>">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Turmas</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strAcao = trim(Request("parAcao"))

if trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))
end if

'Response.write Session("Corte")
'Response.end 

if trim(Request("selFeriado")) <> "" then
	intCDFeriado = trim(Request("selFeriado"))
elseif trim(Request("txtCdFeriado")) <> "0" then
	intCDFeriado = trim(Request("txtCdFeriado"))	
else
	intCDFeriado = ""	
end if
	
strNomeFeriado	= Ucase(trim(Request("txtNomeFeriado")))	
strDtFeriado 	= trim(Request("txtDtFeriado"))
strFeriNacio	= trim(Request("rdFeriNacional"))
'intCdCT			= trim(Request("selCT_Selecionado"))
intCDsSalas		= trim(Request("txtSalas_Selecionadas"))
				
'Response.write "strAcao - " & strAcao & "<br>"		
'Response.write "intCDFeriado - " & intCDFeriado & "<br>"
'Response.write "strNomeFeriado - " & strNomeFeriado & "<br>"
'Response.write "strDtFeriado - " & strDtFeriado & "<br>"
'Response.write "intCdCT - " & intCdCT & "<br>"
'Response.write "intCDsSalas - " & intCDsSalas & "<br>"
'Response.write "strFeriNacio - " & strFeriNacio & "<br><br>"
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Feriado"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Feriado"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Feriado"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE FERIADO ************************************************
if strAcao = "I" then	

	strVerificaFeriado = ""
	strVerificaFeriado = strVerificaFeriado & "SELECT FERI_TX_NOME_FERIADO "
	strVerificaFeriado = strVerificaFeriado & "FROM GRADE_FERIADO "
	strVerificaFeriado = strVerificaFeriado & "WHERE FERI_TX_NOME_FERIADO = '" & strNomeFeriado & "'"		
	'Response.write strVerificaFeriado
	'Response.end
		
	Set rdsVerificaFeriado = db_banco.Execute(strVerificaFeriado)			
	
	if not rdsVerificaFeriado.EOF then
		strMSG = "Já existe feriado cadastrado com o nome - " & rdsVerificaFeriado("FERI_TX_NOME_FERIADO") & "."
	else			
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(FERI_CD_FERIADO) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_FERIADO "			
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCDFeriado = 1
			else
				intCDFeriado = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCDFeriado = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
	
		strSQLIncFeriado = ""
		strSQLIncFeriado = strSQLIncFeriado & "INSERT INTO GRADE_FERIADO (FERI_CD_FERIADO, FERI_TX_NOME_FERIADO, FERI_DT_DATA_FERIADO, "
		strSQLIncFeriado = strSQLIncFeriado & "FERI_TX_TIPO_FERIADO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncFeriado = strSQLIncFeriado & "VALUES(" & intCDFeriado & ",'" & strNomeFeriado & "','" & strDtFeriado & "','" & strFeriNacio & "'," 
		strSQLIncFeriado = strSQLIncFeriado & "'I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLIncFeriado & "<br><br>"
		'Response.end	
		
		'Response.write strFeriNacio & "<br>"
		'Response.end	
		
		on error resume next
			db_banco.Execute(strSQLIncFeriado)	
			
		'*** FERIADO MUNICIPAL - TERÁ QUE GRAVAR EM SALA - FERIADO
		strSQLIncFeriadoSala = ""
		if strFeriNacio = 1 then				
		
			'*** LIMPA OS REGISTROS DA TABELA PARA O FERIADO SELECIONADO **
			strSQLDelFeriadoCT = ""
			strSQLDelFeriadoCT = strSQLDelFeriadoCT & "DELETE FROM GRADE_FERIADO_SALA "	
			strSQLDelFeriadoCT = strSQLDelFeriadoCT & "WHERE FERI_CD_FERIADO = " & intCdFeriado 
			strSQLDelFeriadoCT = strSQLDelFeriadoCT & " AND CORT_CD_CORTE = " & Session("Corte")
			'response.write strSQLDelFeriadoCT
			'Response.end		
			
			db_banco.Execute(strSQLDelFeriadoCT)	
			
			'*** CADASTRA AS NOVOS REGISTROS ***			
			vetCDsSalas = split(intCDsSalas,",")
						
			r = 0			
			for r = lbound(vetCDsSalas) to Ubound(vetCDsSalas)				
				
				'Response.write vetCDsSalas(r) & "<br><br>"		
				
				if vetCDsSalas(r) <> "" then				
					intCdSalaResult = cint(vetCDsSalas(r))
							
					strSQLIncFeriadoSala = ""		
					strSQLIncFeriadoSala = strSQLIncFeriadoSala & "INSERT INTO GRADE_FERIADO_SALA "
					strSQLIncFeriadoSala = strSQLIncFeriadoSala & "(CORT_CD_CORTE, SALA_CD_SALA, FERI_CD_FERIADO, "
					strSQLIncFeriadoSala = strSQLIncFeriadoSala & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
					strSQLIncFeriadoSala = strSQLIncFeriadoSala & "VALUES(" & Session("Corte") & "," & intCdSalaResult & "," & intCDFeriado & "," 
					strSQLIncFeriadoSala = strSQLIncFeriadoSala & "'I','" & Session("CdUsuario") & "',GETDATE())"					
					'response.write strSQLIncFeriadoSala & "<br><br>"
					'Response.end					
					db_banco.Execute(strSQLIncFeriadoSala)
				end if
			next					
		end if				
			
		if err.number = 0 then		
			strMSG = "Feriado foi incluído com sucesso."
		else
			strMSG = "Houve um erro na inclusão do feriado (" & err.description & ")"
		end if	
	end if
	
	rdsVerificaFeriado.close
	set rdsVerificaFeriado = nothing
	
'************************************** ALTERAÇÃO  DE FERIADO ************************************************	
elseif strAcao = "A" then			
		
	'*** LIMPA OS REGISTROS DA TABELA PARA O FERIADO SELECIONADO **
	strSQLDelFeriadoCT = ""
	strSQLDelFeriadoCT = strSQLDelFeriadoCT & "DELETE FROM GRADE_FERIADO_SALA "	
	strSQLDelFeriadoCT = strSQLDelFeriadoCT & "WHERE FERI_CD_FERIADO = " & intCdFeriado 
	strSQLDelFeriadoCT = strSQLDelFeriadoCT & " AND CORT_CD_CORTE = " & Session("Corte")
	'response.write strSQLDelFeriadoCT
	'Response.end		
	
	db_banco.Execute(strSQLDelFeriadoCT)	
				
	'*** ALTERA AS NOVOS REGISTROS ***			
	vetCDsSalas = split(intCDsSalas,",")
				
	r = 0			
	for r = lbound(vetCDsSalas) to Ubound(vetCDsSalas)				
		
		'Response.write vetCDsSalas(r) & "<br><br>"		
		
		if vetCDsSalas(r) <> "" then				
			intCdSalaResult = cint(vetCDsSalas(r))
					
			strSQLIncFeriadoSala = ""		
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "INSERT INTO GRADE_FERIADO_SALA "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "(CORT_CD_CORTE,SALA_CD_SALA, FERI_CD_FERIADO, "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "VALUES(" & Session("Corte") & "," & intCdSalaResult & "," & intCDFeriado & "," 
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "'I','" & Session("CdUsuario") & "',GETDATE())"					
			'response.write strSQLIncFeriadoSala & "<br><br>"
			'Response.end					
			db_banco.Execute(strSQLIncFeriadoSala)
		end if
	next						
		
	strSQLAltFeriado = ""
	strSQLAltFeriado = strSQLAltFeriado & "UPDATE GRADE_FERIADO "
	strSQLAltFeriado = strSQLAltFeriado & "SET FERI_TX_NOME_FERIADO = '" & strNomeFeriado & "',"		
	strSQLAltFeriado = strSQLAltFeriado & "FERI_DT_DATA_FERIADO = '" & strDtFeriado & "',"		
	strSQLAltFeriado = strSQLAltFeriado & "FERI_TX_TIPO_FERIADO = '" & strFeriNacio & "',"		
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltFeriado = strSQLAltFeriado & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltFeriado = strSQLAltFeriado & "WHERE FERI_CD_FERIADO = " & intCdFeriado 	
	'Response.write strSQLAltFeriado
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltFeriado)
					
	if err.number = 0 then		
		strMSG = "Feriado foi alterado com sucesso."
	else
		strMSG = "Houve um erro na alteração do feriado (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE FERIADO ************************************************	
elseif strAcao = "E" then 
				
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmFeriado = ""	
	
	strSQLVeriFeriadoSala = ""			
	strSQLVeriFeriadoSala = strSQLVeriFeriadoSala & "SELECT FERIADO.FERI_CD_FERIADO, FERIADO.FERI_TX_NOME_FERIADO "
	strSQLVeriFeriadoSala = strSQLVeriFeriadoSala & "FROM GRADE_FERIADO FERIADO, GRADE_FERIADO_SALA FERI_SALA "
	strSQLVeriFeriadoSala = strSQLVeriFeriadoSala & "WHERE FERIADO.FERI_CD_FERIADO = FERI_SALA.FERI_CD_FERIADO "
	strSQLVeriFeriadoSala = strSQLVeriFeriadoSala & "AND FERIADO.FERI_CD_FERIADO = " & intCdFeriado 
	strSQLVeriFeriadoSala = strSQLVeriFeriadoSala & " AND FERI_SALA.CORT_CD_CORTE = " & Session("Corte") 	
	'Response.write strSQLVeriFeriadoSala & "<br>"
	'Response.end		
			
	set rsVeriFeriadoSala = db_banco.Execute(strSQLVeriFeriadoSala)		
			
	if not rsVeriFeriadoSala.eof then	
		blnPodeDeletar = false
		strAssociacao = "Sala"
		strNmFeriado = rsVeriFeriadoSala("FERI_TX_NOME_FERIADO")
	end if
				
	'Response.write blnPodeDeletar & " - MSG - " & strAssociacao
	'Response.end					
				
	if blnPodeDeletar then			
		
		'*** LIMPA OS REGISTROS DA TABELA PARA O FERIADO SELECIONADO **
		'strSQLDelFeriadoCT = ""
		'strSQLDelFeriadoCT = strSQLDelFeriadoCT & "DELETE FROM GRADE_FERIADO_SALA "	
		'strSQLDelFeriadoCT = strSQLDelFeriadoCT & "WHERE FERI_CD_FERIADO = " & intCdFeriado 
		'strSQLDelFeriadoCT = strSQLDelFeriadoCT & " AND CORT_CD_CORTE = " & Session("Corte")
		'response.write strSQLDelFeriadoCT
		'Response.end		
		
		'db_banco.Execute(strSQLDelFeriadoCT)				
			
		strSQLDelFeriado = ""
		strSQLDelFeriado = strSQLDelFeriado & "DELETE FROM GRADE_FERIADO "	
		strSQLDelFeriado = strSQLDelFeriado & "WHERE FERI_CD_FERIADO = " & intCdFeriado 	
		'Response.write strSQLDelFeriado
		'Response.end
		
		on error resume next
			db_banco.Execute(strSQLDelFeriado)
		
		if err.number = 0 then		
			strMSG = "Feriado foi excluído com sucesso."
		else
			strMSG = "Houve um erro na exclusão do feriado (" & err.description & ")"
		end if	
	else
		strMSG = "O Feriado - " & strNmFeriado & " não pode ser excluído, pois existe associação do mesmo em " & strAssociacao & "."
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
					<a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"> 
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="35" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
				</tr>
					
				<tr>						
				  <td width="205" height="29"></td>						
				  <td width="93" height="29" valign="middle" align="left"></td>						
				  <td height="29" valign="middle" align="left" width="29"> 				 
				   	<a href="sel_feriado.asp?selCorte=<%=Session("Corte")%>">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Feriados</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strAcao = trim(Request("parAcao"))

if trim(Request("selSala")) <> "" then	
	strSala = trim(Request("selSala"))
elseif trim(Request("txtSala")) <> "" then
	strSala	 = trim(Request("txtSala"))	
else
	strSala	 = ""	
end if
	
if trim(Request("selCT")) <> "" then
	strCdCT = trim(Request("selCT"))
elseif trim(Request("txtNomeSala")) <> "" then
	strCdCT = trim(Request("txtNomeSala"))
else
	strCdCT = ""
end if

strNomeSala		= Ucase(trim(Request("txtNomeSala")))		

'strDiretoria 	= trim(Request("txtDiretoria"))
'strLocal 		= trim(Request("txtLocal"))
strCapacidade	= trim(Request("txtCapacidade"))	
strUniversidade	= trim(Request("pUniv"))
	
'Response.write "strAcao - " & strAcao & "<br>"
'Response.write "strSala - " & strSala & "<br>"
'Response.write "strCdCT - " & strCdCT & "<br>"
'Response.write "strNomeSala - " & strNomeSala & "<br>"	
''Response.write "strLocal - " & strLocal & "<br>"
'Response.write "strCapacidade - " & strCapacidade & "<br>"	
'Response.write "strUniversidade - " & strUniversidade & "<br>"	
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Sala"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Sala"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Sala"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE SALA ************************************************
if strAcao = "I" then	

	strVerificaSala = ""
	strVerificaSala = strVerificaSala & "SELECT SALA_TX_NOME_SALA "
	strVerificaSala = strVerificaSala & "FROM GRADE_SALA "
	strVerificaSala = strVerificaSala & "WHERE SALA_TX_NOME_SALA = '" & strNomeSala & "'"		
	strVerificaSala = strVerificaSala & " AND CORT_CD_CORTE = " & Session("Corte") 
	'Response.write strVerificaSala
	'Response.end
		
	Set rdsVerificaSala = db_banco.Execute(strVerificaSala)			
	
	if not rdsVerificaSala.EOF then
		strMSG = "Já existe sala cadastrada com o nome - " & rdsVerificaSala("SALA_TX_NOME_SALA") & "."
	else			
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(SALA_CD_SALA) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_SALA "	
		strVerificaCod = strVerificaCod & "WHERE CORT_CD_CORTE = " & Session("Corte") 		
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCDSala = 1
			else				
				intCDSala = rdsVerificaCod("COD_MAIOR") + 1
			end if			
		else
			intCDSala = 1
		end if
						
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
	
		strSQLIncSala = ""
		strSQLIncSala = strSQLIncSala & "INSERT INTO GRADE_SALA (CORT_CD_CORTE, SALA_CD_SALA, CTRO_CD_CENTRO_TREINAMENTO, SALA_TX_NOME_SALA, "
		strSQLIncSala = strSQLIncSala & "SALA_NUM_CAPACIDADE, SALA_CD_UC, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncSala = strSQLIncSala & "VALUES(" & Session("Corte") & "," & intCDSala & "," & strCdCT & ",'" & strNomeSala & "','" & strCapacidade & "','" 
		strSQLIncSala = strSQLIncSala & strUniversidade & "','I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLIncSala
		'Response.end	
		
		on error resume next
			db_banco.Execute(strSQLIncSala)	

		if err.number = 0 then		
			strMSG = "Sala foi incluída com sucesso."
		else
			strMSG = "Houve um erro na inclusão da sala (" & err.description & ")."
		end if	
	end if
	
	rdsVerificaSala.close
	set rdsVerificaSala = nothing
	
'************************************** ALTERAÇÃO  DE SALA ************************************************	
elseif strAcao = "A" then			
		
	strSQLAltSala = ""
	strSQLAltSala = strSQLAltSala & "UPDATE GRADE_SALA "
	strSQLAltSala = strSQLAltSala & "SET SALA_TX_NOME_SALA = '" & strNomeSala & "',"	
	strSQLAltSala = strSQLAltSala & "CTRO_CD_CENTRO_TREINAMENTO = " & strCdCT & ","	
	strSQLAltSala = strSQLAltSala & "SALA_NUM_CAPACIDADE = " & strCapacidade & ","		
	strSQLAltSala = strSQLAltSala & "SALA_CD_UC = '" & strUniversidade & "',"		
	strSQLAltSala = strSQLAltSala & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltSala = strSQLAltSala & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltSala = strSQLAltSala & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltSala = strSQLAltSala & "WHERE SALA_CD_SALA = " & strSala	
	strSQLAltSala = strSQLAltSala & " AND CORT_CD_CORTE = " & Session("Corte") 	
	'Response.write strSQLAltSala
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltSala)
					
	if err.number = 0 then		
		strMSG = "Sala foi alterada com sucesso."
	else
		strMSG = "Houve um erro na alteração da sala (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE SALA ************************************************	
elseif strAcao = "E" then 
				
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmSala = ""		
		
	strSQLVeriSalaTurma = ""			
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & "SELECT SALA.SALA_CD_SALA, SALA.SALA_TX_NOME_SALA "
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & "FROM GRADE_SALA SALA, GRADE_TURMA TURMA "
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & "WHERE TURMA.SALA_CD_SALA = SALA.SALA_CD_SALA "
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & "AND SALA.SALA_CD_SALA = " & strSala 
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & " AND SALA.CORT_CD_CORTE = " & Session("Corte") 	
	strSQLVeriSalaTurma = strSQLVeriSalaTurma & " AND TURMA.CORT_CD_CORTE = " & Session("Corte") 	
	'Response.write strSQLVeriSalaTurma & "<br>"
	'Response.end		
			
	set rsVeriSalaTurma = db_banco.Execute(strSQLVeriSalaTurma)		
			
	if not rsVeriSalaTurma.eof then	
		blnPodeDeletar = false
		strAssociacao = "Turma"
		strNmSala = rsVeriSalaTurma("SALA_TX_NOME_SALA")
	end if
			
	strSQLSalaFeriado = ""			
	strSQLSalaFeriado = strSQLSalaFeriado & "SELECT SALA.SALA_CD_SALA, SALA.SALA_TX_NOME_SALA "
	strSQLSalaFeriado = strSQLSalaFeriado & "FROM GRADE_SALA SALA, GRADE_FERIADO_SALA SALA_FERI "
	strSQLSalaFeriado = strSQLSalaFeriado & "WHERE SALA_FERI.SALA_CD_SALA = SALA.SALA_CD_SALA "
	strSQLSalaFeriado = strSQLSalaFeriado & "AND SALA.SALA_CD_SALA = " & strSala 
	strSQLSalaFeriado = strSQLSalaFeriado & " AND SALA.CORT_CD_CORTE = " & Session("Corte") 	
	strSQLSalaFeriado = strSQLSalaFeriado & " AND SALA_FERI.CORT_CD_CORTE = " & Session("Corte") 	
	'Response.write strSQLSalaFeriado & "<br>"
	'Response.end		
			
	set rsVeriSalaFeriado = db_banco.Execute(strSQLSalaFeriado)		
			
	if not rsVeriSalaFeriado.eof then	
		blnPodeDeletar = false
		if strAssociacao <> "" then
			strAssociacao = strAssociacao & " e Feriado"
		else
			strAssociacao = "Feriado"
		end if
		strNmSala = rsVeriSalaFeriado("SALA_TX_NOME_SALA")
	end if		
	
	'Response.write blnPodeDeletar & " - MSG - " & strAssociacao
	'Response.end		
	if blnPodeDeletar then
		strSQLDelSala = ""
		strSQLDelSala = strSQLDelSala & "DELETE FROM GRADE_SALA "	
		strSQLDelSala = strSQLDelSala & "WHERE SALA_CD_SALA = " & strSala 	
		strSQLDelSala = strSQLDelSala & " AND CORT_CD_CORTE = " & Session("Corte") 	
		'Response.write strSQLDelSala
		'Response.end
		
		on error resume next
			db_banco.Execute(strSQLDelSala)
		
		if err.number = 0 then		
			strMSG = "Sala foi excluída com sucesso."
		else
			strMSG = "Houve um erro na exclusão da sala (" & err.description & ")"
		end if	
	else
		strMSG = "A Sala - " & strNmSala & " não pode ser excluída, pois existe associação da mesma em " & strAssociacao & "."
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
				   	<a href="sel_sala.asp">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Salas</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strAcao = trim(Request("parAcao"))

Session("Corte") = trim(Request("selCorte"))

if trim(Request("selCT")) <> "" then
	intCdCT = trim(Request("selCT"))
elseif trim(Request("txtCdCT")) <> "0" then
	intCdCT = trim(Request("txtCdCT"))
else
	intCdCT = ""	
end if
	
strNomeCT	= Ucase(trim(Request("txtNomeCT")))	
intCdLocalidade	= Ucase(trim(Request("selLocalidade")))	
		
'Response.write "strAcao - " & strAcao & "<br>"		
'Response.write "intCdCT - " & intCdCT & "<br>"
'Response.write "strNomeCT - " & strNomeCT & "<br>"
'Response.write "intCdLocalidade - " & intCdLocalidade & "<br>"
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Centro de Treinamento"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Centro de Treinamento"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Centro de Treinamento"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE CENTRO DE TREINAMENTO ************************************************
if strAcao = "I" then	

	strVerificaCT = ""
	strVerificaCT = strVerificaCT & "SELECT CTRO_TX_NOME_CENTRO_TREINAMENTO "
	strVerificaCT = strVerificaCT & "FROM GRADE_CENTRO_TREINAMENTO "
	strVerificaCT = strVerificaCT & "WHERE CTRO_TX_NOME_CENTRO_TREINAMENTO = '" & strNomeCT & "'"	
	strVerificaCT = strVerificaCT & " AND CORT_CD_CORTE=" & Session("Corte") 		
	'Response.write strVerificaCT
	'Response.end
		
	Set rdsVerificaCT = db_banco.Execute(strVerificaCT)			
	
	if not rdsVerificaCT.EOF then
		strMSG = "Já existe centro de treinamento cadastrado com o nome - " & rdsVerificaCT("CTRO_TX_NOME_CENTRO_TREINAMENTO") & "."
	else			
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(CTRO_CD_CENTRO_TREINAMENTO) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_CENTRO_TREINAMENTO"			
		strVerificaCod = strVerificaCod & " WHERE CORT_CD_CORTE=" & Session("Corte") 
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCdCT = 1
			else
				intCdCT = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCdCT = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
	
		strSQLCT = ""
		strSQLCT = strSQLCT & "INSERT INTO GRADE_CENTRO_TREINAMENTO (CORT_CD_CORTE, CTRO_CD_CENTRO_TREINAMENTO, LOC_CD_LOCALIDADE, CTRO_TX_NOME_CENTRO_TREINAMENTO, "
		strSQLCT = strSQLCT & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLCT = strSQLCT & "VALUES(" & Session("Corte") & "," & intCdCT & "," & intCdLocalidade & ",'" & strNomeCT & "'," 
		strSQLCT = strSQLCT & "'I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLCT
		'Response.end	
		
		on error resume next
			db_banco.Execute(strSQLCT)	

		if err.number = 0 then		
			strMSG = "Centro de treinamento foi incluído com sucesso."
		else
			strMSG = "Houve um erro na inclusão do centro de treinamento (" & err.description & ")"
		end if	
	end if
	
	rdsVerificaCT.close
	set rdsVerificaCT = nothing
'************************************** ALTERAÇÃO DE CENTRO DE TREINAMENTO ************************************************	
elseif strAcao = "A" then			
		
	strSQLAltCT = ""
	strSQLAltCT = strSQLAltCT & "UPDATE GRADE_CENTRO_TREINAMENTO "
	strSQLAltCT = strSQLAltCT & "SET CTRO_TX_NOME_CENTRO_TREINAMENTO = '" & strNomeCT & "',"		
	strSQLAltCT = strSQLAltCT & "LOC_CD_LOCALIDADE = " & intCdLocalidade & ","
	strSQLAltCT = strSQLAltCT & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltCT = strSQLAltCT & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltCT = strSQLAltCT & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltCT = strSQLAltCT & "WHERE CTRO_CD_CENTRO_TREINAMENTO = " & intCdCT 
	strSQLAltCT = strSQLAltCT & " AND CORT_CD_CORTE = " & Session("Corte")
	'Response.write strSQLAltCT
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltCT)
					
	if err.number = 0 then		
		strMSG = "Centro de treinamento foi alterado com sucesso."
	else
		strMSG = "Houve um erro na alteração do centro de treinamento (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE CENTRO DE TREINAMENTO ************************************************	
elseif strAcao = "E" then 
					
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmCT = ""	
	
	'*** VERIFICA EM ORGÃO MENOR ****
	strSQLVeriSala = ""			
	strSQLVeriSala = strSQLVeriSala & "SELECT CT.CTRO_CD_CENTRO_TREINAMENTO, CT.CTRO_TX_NOME_CENTRO_TREINAMENTO "
	strSQLVeriSala = strSQLVeriSala & "FROM GRADE_SALA SALA, GRADE_CENTRO_TREINAMENTO CT "
	strSQLVeriSala = strSQLVeriSala & "WHERE SALA.CTRO_CD_CENTRO_TREINAMENTO = CT.CTRO_CD_CENTRO_TREINAMENTO "
	strSQLVeriSala = strSQLVeriSala & "AND  CT.CTRO_CD_CENTRO_TREINAMENTO = " & intCdCT 
	strSQLVeriSala = strSQLVeriSala & " AND CT.CORT_CD_CORTE = " & Session("Corte") 
	strSQLVeriSala = strSQLVeriSala & " AND SALA.CORT_CD_CORTE = " & Session("Corte") 		
	'Response.write strSQLVeriSala & "<br>"
	'Response.end		
			
	set rsVeriSala = db_banco.Execute(strSQLVeriSala)		
			
	if not rsVeriSala.eof then	
		blnPodeDeletar = false
		strAssociacao = "Sala"
		strNmCT = rsVeriSala("CTRO_TX_NOME_CENTRO_TREINAMENTO")
	end if				
	
	if blnPodeDeletar then		
								
		strSQLDelCT = ""
		strSQLDelCT = strSQLDelCT & "DELETE FROM GRADE_CENTRO_TREINAMENTO "	
		strSQLDelCT = strSQLDelCT & "WHERE CTRO_CD_CENTRO_TREINAMENTO = " & intCdCT 
		strSQLDelCT = strSQLDelCT & " AND CORT_CD_CORTE = " & Session("Corte")	
		'Response.write strSQLDelCT
		'Response.end
		
		on error resume next			
			db_banco.Execute(strSQLDelCT)
		
		if err.number = 0 then		
			strMSG = "Centro de treinamento foi excluído com sucesso."
		else
			strMSG = "Houve um erro na exclusão do centro de treinamento (" & err.description & ")"
		end if	
	else
		strMSG = "O Centro de Treinamento - " & strNmCT & " não pode ser excluído, pois existe associação do mesmo em " & strAssociacao & "."
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
					<a href="Inclui_altera_turma.asp?selSala=<%=strSala%>"> 
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="35" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
				</tr>
					
				<tr>						
				  <td width="205" height="29"></td>						
				  <td width="93" height="29" valign="middle" align="left"></td>						
				  <td height="29" valign="middle" align="left" width="29"> 				 
				   	<a href="sel_CT.asp">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Centro de Treinamento</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
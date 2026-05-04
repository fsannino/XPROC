<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strAcao = trim(Request("parAcao"))

if trim(Request("selLocalidade")) <> "" then
	intCDLocalidade = trim(Request("selLocalidade"))
elseif trim(Request("txtCdLocalidade")) <> "0" then
	intCDLocalidade = trim(Request("txtCdLocalidade"))	
else
	intCDLocalidade = ""	
end if
	
strNomeLocalidade	= Ucase(trim(Request("txtNomeLocalidade")))	
		
'Response.write "strAcao - " & strAcao & "<br>"		
'Response.write "intCDLocalidade - " & intCDLocalidade & "<br>"
'Response.write "strNomeLocalidade - " & strNomeLocalidade & "<br>"
'response.end 

if strAcao = "I" then
	strNomeAcao = "Inclusão de Localidade"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Localidade"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Localidade"
end if
					
strMSG =  ""				
'************************************** INCLUSÃO DE LOCALIDADE ************************************************
if strAcao = "I" then	

	strVerificaLocalidade = ""
	strVerificaLocalidade = strVerificaLocalidade & "SELECT LOC_TX_NOME_LOCALIDADE "
	strVerificaLocalidade = strVerificaLocalidade & "FROM GRADE_LOCALIDADE "
	strVerificaLocalidade = strVerificaLocalidade & "WHERE LOC_TX_NOME_LOCALIDADE = '" & strNomeLocalidade & "'"		
	'Response.write strVerificaLocalidade
	'Response.end
		
	Set rdsVerificaLocalidade = db_banco.Execute(strVerificaLocalidade)			
	
	if not rdsVerificaLocalidade.EOF then
		strMSG = "Já existe localidade cadastrada com o nome - " & rdsVerificaLocalidade("LOC_TX_NOME_LOCALIDADE") & "."
	else			
		strVerificaCod = ""
		strVerificaCod = strVerificaCod & "SELECT MAX(LOC_CD_LOCALIDADE) as COD_MAIOR "
		strVerificaCod = strVerificaCod & "FROM GRADE_LOCALIDADE"			
		'Response.write strVerificaCod
		'Response.end
		Set rdsVerificaCod = db_banco.Execute(strVerificaCod)		
		
		if not rdsVerificaCod.eof then
			if isnull(rdsVerificaCod("COD_MAIOR")) then
				intCDLocalidade = 1
			else
				intCDLocalidade = rdsVerificaCod("COD_MAIOR") + 1
			end if
		else
			intCDLocalidade = 1
		end if
		
		rdsVerificaCod.close
		set rdsVerificaCod = nothing
	
		strSQLIncLocalidade = ""
		strSQLIncLocalidade = strSQLIncLocalidade & "INSERT INTO GRADE_LOCALIDADE (LOC_CD_LOCALIDADE, LOC_TX_NOME_LOCALIDADE, "
		strSQLIncLocalidade = strSQLIncLocalidade & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncLocalidade = strSQLIncLocalidade & "VALUES(" & intCDLocalidade & ",'" & strNomeLocalidade & "'," 
		strSQLIncLocalidade = strSQLIncLocalidade & "'I','" & Session("CdUsuario") & "',GETDATE())"	
		'response.write strSQLIncLocalidade
		'Response.end	
		
		on error resume next
			db_banco.Execute(strSQLIncLocalidade)	

		if err.number = 0 then		
			strMSG = "Localidade foi incluída com sucesso."
		else
			strMSG = "Houve um erro na inclusão da localidade (" & err.description & ")"
		end if	
	end if
	
	rdsVerificaLocalidade.close
	set rdsVerificaLocalidade = nothing
	
'************************************** ALTERAÇÃO  DE LOCALIDADE ************************************************	
elseif strAcao = "A" then			
		
	strSQLAltLocalidade = ""
	strSQLAltLocalidade = strSQLAltLocalidade & "UPDATE GRADE_LOCALIDADE "
	strSQLAltLocalidade = strSQLAltLocalidade & "SET LOC_TX_NOME_LOCALIDADE = '" & strNomeLocalidade & "',"		
	strSQLAltLocalidade = strSQLAltLocalidade & "ATUA_TX_OPERACAO = 'A'," 
	strSQLAltLocalidade = strSQLAltLocalidade & "ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'," 
	strSQLAltLocalidade = strSQLAltLocalidade & "ATUA_DT_ATUALIZACAO = GETDATE() "	
	strSQLAltLocalidade = strSQLAltLocalidade & "WHERE LOC_CD_LOCALIDADE = " & intCDLocalidade 	
	'Response.write strSQLAltLocalidade
	'Response.end			
	
	on error resume next
		db_banco.Execute(strSQLAltLocalidade)
					
	if err.number = 0 then		
		strMSG = "Localidade foi alterada com sucesso."
	else
		strMSG = "Houve um erro na alteração da localidade (" & err.description & ")"
	end if			
	
'************************************** EXCLUSÃO DE LOCALIDADE ************************************************	
elseif strAcao = "E" then 
				
	blnPodeDeletar = True			
	strAssociacao = ""			
	strNmLocalidade = ""	
	
	'*** VERIFICA EM ORGÃO MENOR ****
	strSQLVeriCT = ""			
	strSQLVeriCT = strSQLVeriCT & "SELECT LOCAL.LOC_CD_LOCALIDADE, LOCAL.LOC_TX_NOME_LOCALIDADE "
	strSQLVeriCT = strSQLVeriCT & "FROM GRADE_LOCALIDADE LOCAL, GRADE_CENTRO_TREINAMENTO CT "
	strSQLVeriCT = strSQLVeriCT & "WHERE LOCAL.LOC_CD_LOCALIDADE = CT.LOC_CD_LOCALIDADE "
	strSQLVeriCT = strSQLVeriCT & "AND LOCAL.LOC_CD_LOCALIDADE = " & intCDLocalidade 
	'Response.write strSQLVeriCT & "<br>"
	'Response.end		
			
	set rsVeriCT = db_banco.Execute(strSQLVeriCT)		
			
	if not rsVeriCT.eof then	
		blnPodeDeletar = false
		strAssociacao = "Centro de Treinamento"
		strNmLocalidade = rsVeriCT("LOC_TX_NOME_LOCALIDADE")
	end if				
	
	if blnPodeDeletar then		
				
		strSQLDelLocalidade = ""
		strSQLDelLocalidade = strSQLDelLocalidade & "DELETE FROM GRADE_LOCALIDADE "	
		strSQLDelLocalidade = strSQLDelLocalidade & "WHERE LOC_CD_LOCALIDADE = " & intCDLocalidade 	
		'Response.write strSQLDelLocalidade
		'Response.end
		
		on error resume next
			db_banco.Execute(strSQLDelLocalidade)
		
		if err.number = 0 then		
			strMSG = "Localidade foi excluída com sucesso."
		else
			strMSG = "Houve um erro na exclusão da localidade (" & err.description & ")"
		end if
	else
		strMSG = "A Localidade - " & strNmLocalidade & " não pode ser excluída, pois existe associação da mesma em " & strAssociacao & "."
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
				   	<a href="sel_localidade.asp">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Localidade</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
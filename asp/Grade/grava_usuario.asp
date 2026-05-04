<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

strAcao 			= Request("pAcao")
strCdUsuario 		= Ucase(Request("txtUsuarioAcesso"))
strNomeUsuario 		= Ucase(Request("pNomeUsua"))
strCategoria 		= Request("rdbCategoria")

if strAcao = "I" then
	strNomeAcao = "Inclusão de Usuário"
elseif strAcao = "A" then
	strNomeAcao = "Alteração de Usuário"
elseif strAcao = "E" then
	strNomeAcao = "Exclusão de Usuário"
end if

'response.write "<br><br><br> strAcao " & strAcao & "<br>"
'response.write "strCdUsuario " & strCdUsuario  & "<br>"
'response.write "strNomeUsuario " & strNomeUsuario  & "<br>"
'response.write "strCategoria " & strCategoria  & "<br>"
'response.end

if strAcao = "I" then

	sqlVerificaUsuario = ""
	sqlVerificaUsuario = sqlVerificaUsuario & "SELECT USUA_TX_NOME_USUARIO " 
	sqlVerificaUsuario = sqlVerificaUsuario & "FROM GRADE_USUARIO "
	sqlVerificaUsuario = sqlVerificaUsuario & "WHERE USUA_CD_USUARIO = '" & strCdUsuario & "'"
	'Response.write sqlVerificaUsuario
	'Response.end
	Set rdsVerificaUsuario = db_Cogest.Execute(sqlVerificaUsuario)			
		
	if not rdsVerificaUsuario.EOF then
		strMSG = "Já existe usuário cadastrado com o nome - " & rdsVerificaUsuario("USUA_TX_NOME_USUARIO") & "."
	else			
		on error resume next			
			sqlNovoUsuario = ""
			sqlNovoUsuario = " INSERT INTO GRADE_USUARIO(USUA_CD_USUARIO, USUA_TX_NOME_USUARIO, "
			sqlNovoUsuario = sqlNovoUsuario & "USUA_TX_CATEGORIA, ATUA_TX_OPERACAO, "
			sqlNovoUsuario = sqlNovoUsuario & "ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
			sqlNovoUsuario = sqlNovoUsuario & ")VALUES('" & strCdUsuario & "','" & strNomeUsuario & "','" & strCategoria & "',"
			sqlNovoUsuario = sqlNovoUsuario & "'I','" & Session("CdUsuario") & "',GETDATE())"
			'Response.write sqlNovoUsuario & "<br><br>"
			'Response.end							
			db_Cogest.Execute(sqlNovoUsuario)
							
			if err.number = 0 then		
				strMSG = "Usuário incluído com sucesso."
			else
				strMSG = "Houve um erro na inclusão de Usuário (" & err.description & ")."
			end if	
	end if	
	
	rdsVerificaUsuario.close
	set rdsVerificaUsuario = nothing
	
elseif strAcao = "A" then
	
	on error resume next			
		sqlAltUsuario = ""
		sqlAltUsuario = "UPDATE GRADE_USUARIO SET"
		sqlAltUsuario = sqlAltUsuario & " USUA_TX_NOME_USUARIO ='" & strNomeUsuario & "'"
		sqlAltUsuario = sqlAltUsuario & ",USUA_TX_CATEGORIA ='" & strCategoria & "'"
		sqlAltUsuario = sqlAltUsuario & ",ATUA_TX_OPERACAO ='A'" 
		sqlAltUsuario = sqlAltUsuario & ",ATUA_CD_NR_USUARIO ='" & Session("CdUsuario") & "'"
		sqlAltUsuario = sqlAltUsuario & ",ATUA_DT_ATUALIZACAO = GETDATE()"
		sqlAltUsuario = sqlAltUsuario & " WHERE USUA_CD_USUARIO = '" & strCdUsuario & "'"			
		'Response.write sqlAltUsuario & "<br><br>"
		'Response.end							
		db_Cogest.Execute(sqlAltUsuario)
						
		if err.number = 0 then		
			strMSG = "Usuário alterado com sucesso."
		else
			strMSG = "Houve um erro na alteração do Usuário (" & err.description & ")."
		end if	
		
elseif strAcao = "E" then
			
	sqlExcUsuario = ""
	sqlExcUsuario = sqlExcUsuario & " DELETE GRADE_USUARIO"		
	sqlExcUsuario = sqlExcUsuario & " WHERE USUA_CD_USUARIO = '" & strCdUsuario & "'"		
	'Response.write sqlExcUsuario & "<br><br>"
	'Response.end															
																							
	on error resume next
		db_Cogest.Execute(sqlExcUsuario)	
			
	if err.number = 0 then		
		strMSG = "Usuário excluído com sucesso."
	else
		strMSG = "Houve um erro na exclusão do Usuário (" & err.description & ")."
	end if		
end if
%>

<html>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
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
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					 <td width="28"></td>  
						<td width="250"></td>
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
			
	  <td width="117" height="1"></td>
			
	  <td width="53" height="1" valign="middle" align="left"></td>
			
	  <td height="1" valign="middle" align="left" width="32"> 
		<a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>">
		<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>				
	  <td height="1" valign="middle" align="left" width="629"> 
		<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
		  </tr>			  
		  <tr>				
			  <td width="117" height="1"></td>						
			  <td width="53" height="1" valign="middle" align="left"></td>						
			  <td height="1" valign="middle" align="left" width="32">					
				<a href="cadastra_usuario.asp">
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
			  <td height="1" valign="middle" align="left" width="629"> 
				<font face="Verdana" color="#330099" size="2">Retornar para Tela de Usuário</font>
			  </td>
		  </tr>		
		  <tr>					
		  <td width="117" height="1"></td>					
		  <td width="53" height="1" valign="middle" align="left"></td>					
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
		  </tr>
		</table>	
  	<%
	db_Cogest.close
	set db_Cogest = nothing
	%>
  <p>&nbsp;</p>

	</body>
</html>


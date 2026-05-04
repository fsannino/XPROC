<%
Response.Expires = 0

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strCdCurso = trim(Request("hdCurso"))

strDesc = trim(Request("rdDescentralizado"))
if strDesc = "0" then
	srtNomeDesc = "CENTRALIZADO"
else
	srtNomeDesc = "DESCENTRALIZADO"
end if

strInLoco = trim(Request("rdInLoco"))
if strInLoco = "0" then
	strNomeInLoco = "S"
else
	strNomeInLoco = "N"
end if

intCDsUnidade = trim(Request("txtUnid_Selecionadas"))

'Response.write "strCdCurso - " & strCdCurso & "<br>"
'Response.write "strDesc - " & srtNomeDesc & "<br>"
'Response.write "strInLoco - " & strNomeInLoco & "<br><br>"
'Response.write "intCDsUnidade - " & intCDsUnidade & "<br><br>"
'response.end 

strMSG =  ""			
	
'************************************** ATUALIZAÇÃO DE CURSO ************************************************	
strSQLAtuaCurso = ""
strSQLAtuaCurso = strSQLAtuaCurso & "UPDATE GRADE_CURSO SET "
strSQLAtuaCurso = strSQLAtuaCurso & "CURS_TX_CENTRALIZADO = '" & srtNomeDesc & "',"		
strSQLAtuaCurso = strSQLAtuaCurso & "CURS_TX_IN_LOCO= '" & strNomeInLoco & "' " 	
strSQLAtuaCurso = strSQLAtuaCurso & "WHERE CURS_CD_CURSO = '" & strCdCurso	& "' "
strSQLAtuaCurso = strSQLAtuaCurso & "AND CORT_CD_CORTE = " & Session("Corte")
'Response.write strSQLAtuaCurso
'Response.end			

'*** LIMPA OS REGISTROS DA TABELA PARA O FERIADO SELECIONADO **
strSQLDelUnidCurso = ""
strSQLDelUnidCurso = strSQLDelUnidCurso & "DELETE FROM GRADE_CURSO_UNIDADE "	
strSQLDelUnidCurso = strSQLDelUnidCurso & "WHERE CURS_CD_CURSO = '" & strCdCurso & "'"
strSQLDelUnidCurso = strSQLDelUnidCurso & " AND CORT_CD_CORTE=" & Session("Corte")
'response.write strSQLDelUnidCurso
'Response.end		

db_banco.Execute(strSQLDelUnidCurso)	
			
'*** ALTERA AS NOVOS REGISTROS ***			
vetCDsUnidades = split(intCDsUnidade,",")
			
r = 0			
for r = lbound(vetCDsUnidades) to Ubound(vetCDsUnidades)				
		
	if vetCDsUnidades(r) <> "" then				
		intCDUnidResult = vetCDsUnidades(r)
				
		strSQLIncUnidCurso = ""		
		strSQLIncUnidCurso = strSQLIncUnidCurso & "INSERT INTO GRADE_CURSO_UNIDADE "
		strSQLIncUnidCurso = strSQLIncUnidCurso & "(CORT_CD_CORTE, CURS_CD_CURSO, UNID_CD_UNIDADE, "
		strSQLIncUnidCurso = strSQLIncUnidCurso & "ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
		strSQLIncUnidCurso = strSQLIncUnidCurso & "VALUES(" & Session("Corte") & ",'" & strCdCurso & "','" & intCDUnidResult & "'," 
		strSQLIncUnidCurso = strSQLIncUnidCurso & "'I','" & Session("CdUsuario") & "',GETDATE())"					
		'response.write strSQLIncUnidCurso & "<br><br>"
		'Response.end					
		db_banco.Execute(strSQLIncUnidCurso)
	end if
next						
	
on error resume next
	db_banco.Execute(strSQLAtuaCurso)
				
if err.number = 0 then		
	strMSG = "Curso foi atualizado com sucesso."
else
	strMSG = "Houve um erro na atualização do curso (" & err.description & ")"
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Atualização de Curso - Grade de Treinamento</b></font></div>
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
				   	<a href="sel_Curso.asp">			   					
					<img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
						
				  <td height="29" valign="middle" align="left" width="504"> 
					<font face="Verdana" color="#330099" size="2">Retornar para a Tela de Cadastramento de Curso</font></td>
			    </tr>					
		  </table>
    </form>	
	</body>
	<%
	db_banco.close
	set db_banco = nothing
	%>
</html>
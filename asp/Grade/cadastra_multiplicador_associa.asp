<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
'strAcao = trim(Request("pAcao"))
strMostraComboMultTrue = false
strMostraInformMultTrue = false

if trim(Request("selCorte")) <> "" then
	session("Corte") = Request("selCorte")
	intCdCorte = session("Corte")
elseif trim(session("Corte")) <> "" then
	intCdCorte = session("Corte")
end if
			
'Response.write intCdCorte
'Response.end			
			
strNomeAcao ="Associar Multiplicador Extra"
	
if trim(Request("selMultiplicador")) <> "" and trim(Request("selMultiplicador")) <> "0" then
	strVetMult = split(trim(Request("selMultiplicador")),"|")
	intCdMultiplicador = strVetMult(0)
	intTipoMultiplicador = strVetMult(1)
end if	

'******** CORTE ********************
strSQLCorte = ""
strSQLCorte = strSQLCorte & "SELECT CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & intCdCorte
'Response.write strSQLCorte & "<br>"
'Response.end
set rsCorte = db_banco.Execute(strSQLCorte)		

strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")

rsCorte.close
set rsCorte = nothing	
							
'**************** MULTIPLICADOR ****************	
strSQLAltMultiplicador = ""
strSQLAltMultiplicador = strSQLAltMultiplicador & "SELECT CORT_CD_CORTE, MULT_NR_CD_ID_MULT, MULT_TX_NOME_MULTIPLICADOR, ORME_CD_ORG_MENOR, "
strSQLAltMultiplicador = strSQLAltMultiplicador & "MULT_TX_TIPO_MULTIPLICADOR, MULT_TX_RESTRICAO_VIAGEM, MULT_NR_TIPO_MULTIPLICADOR "
strSQLAltMultiplicador = strSQLAltMultiplicador & "FROM GRADE_MULTIPLICADOR " 
strSQLAltMultiplicador = strSQLAltMultiplicador & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador
strSQLAltMultiplicador = strSQLAltMultiplicador & " AND CORT_CD_CORTE = " & intCdCorte 
strSQLAltMultiplicador = strSQLAltMultiplicador & " AND MULT_NR_TIPO_MULTIPLICADOR = " & intTipoMultiplicador
'Response.write strSQLAltMultiplicador & "<br><br>"
'Response.end

Set rdsAltMultiplicador = db_banco.Execute(strSQLAltMultiplicador)			

if not rdsAltMultiplicador.EOF then			

	strNomeMultiplicador	= trim(rdsAltMultiplicador("MULT_TX_NOME_MULTIPLICADOR"))		
	intCdDiretoria		 	= rdsAltMultiplicador("ORME_CD_ORG_MENOR")					
	
	'************ DIRETORIA ****************
	strSQLDiretoria =  ""
	strSQLDiretoria = strSQLDiretoria & "SELECT DIRE_TX_DESC_DIRETORIA "
	strSQLDiretoria = strSQLDiretoria & "FROM GRADE_DIRETORIA "
	strSQLDiretoria = strSQLDiretoria & "WHERE ORLO_CD_ORG_LOT = " & intCdDiretoria
	'Response.WRITE strSQLDiretoria & "<br><br>"
	'Response.END
	
	set rdsDiretoria = db_banco.execute(strSQLDiretoria)
	
	if not rdsDiretoria.eof then
		strNomeDiretoria = rdsDiretoria("DIRE_TX_DESC_DIRETORIA")
	else
		strNomeDiretoria = "N/A"
	end if

	rdsDiretoria.close
	set rdsDiretoria = nothing	

	'*** MULTIPLICADOR X CURSO *****				
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "SELECT DISTINCT MULT_CURSO.CURS_CD_CURSO "
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "AND MULT_CURSO.MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte 
	strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & " AND MULT.CORT_CD_CORTE = " & intCdCorte
	'Response.write strNomeMultiplicadorCurso
	'Response.end

	Set rdsAltMultiplicadorCurso = db_banco.Execute(strNomeMultiplicadorCurso)	
								
	if not rdsAltMultiplicadorCurso.EOF then					
		
		strCountCurso = 0			
		strTodosCursos = ""	
		strTodosCursosQuery = ""	
		
		do while not rdsAltMultiplicadorCurso.eof					
			if strCountCurso = 0 then
				strTodosCursos = rdsAltMultiplicadorCurso("CURS_CD_CURSO")
				strTodosCursosQuery = "'" & rdsAltMultiplicadorCurso("CURS_CD_CURSO") & "'"
			else
				strTodosCursos = strTodosCursos & " - " & rdsAltMultiplicadorCurso("CURS_CD_CURSO")
				strTodosCursosQuery = strTodosCursosQuery & ",'" & rdsAltMultiplicadorCurso("CURS_CD_CURSO") & "'"
			end if			
			
			strCountCurso = strCountCurso + 1
			rdsAltMultiplicadorCurso.movenext
		loop		
	else				
		strTodosCursos = ""		
		strTodosCursosQuery = ""				
	end if	
		
	'*** MULTIPLICADORES QUE POSSUEM CURSO IGUAL AO MULTIPLICADOR EXTRA ***
	strSQLMultiplicador = ""
	strSQLMultiplicador = strSQLMultiplicador & "SELECT DISTINCT MULT.MULT_NR_CD_ID_MULT, MULT.MULT_TX_NOME_MULTIPLICADOR, MULT_NR_TIPO_MULTIPLICADOR "
	strSQLMultiplicador = strSQLMultiplicador & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
	strSQLMultiplicador = strSQLMultiplicador & "WHERE MULT.CORT_CD_CORTE = MULT_CURSO.CORT_CD_CORTE "
	strSQLMultiplicador = strSQLMultiplicador & "AND MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
	strSQLMultiplicador = strSQLMultiplicador & "AND MULT.CORT_CD_CORTE = "& intCdCorte
	strSQLMultiplicador = strSQLMultiplicador & " AND MULT_CURSO.CORT_CD_CORTE = "& intCdCorte
	strSQLMultiplicador = strSQLMultiplicador & " AND MULT.MULT_NR_TIPO_MULTIPLICADOR <> 3 "
	if strTodosCursosQuery <> "" then
		strSQLMultiplicador = strSQLMultiplicador & "AND MULT_CURSO.CURS_CD_CURSO IN (" & strTodosCursosQuery & ") "
	end if
	strSQLMultiplicador = strSQLMultiplicador & "ORDER BY MULT.MULT_TX_NOME_MULTIPLICADOR "	
	'Response.write strSQLMultiplicador & "<br><br>"
	'Response.end
	set rsMultiplicador = db_banco.Execute(strSQLMultiplicador)			
	
	if not rsMultiplicador.eof then
		strMostraComboMultTrue = true
	end if	
end if

rdsAltMultiplicador.close
set rdsAltMultiplicador = nothing	
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript" src="../js/troca_lista.js"></script>
		
		<script language="javascript">
					
			function Confirma()
			{			
				if(document.frmCadMultiplicador.selMultiplicadorTrue.selectedIndex == 0)
				{
					alert("Selecione um MULTIPLICADOR!");
					document.frmCadMultiplicador.selMultiplicadorTrue.focus();
					return;
				}													
																										
				document.frmCadMultiplicador.action="grava_multiplicador.asp";	
				document.frmCadMultiplicador.submit();			
			}	
			
			function submet_pagina(strMultTrue, strCorte)
			{
				var strMult = document.frmCadMultiplicador.selMultiplicador.value;
			
				if (document.frmCadMultiplicador.selMultiplicadorTrue.selectedIndex != 0)
				{
					window.location.href = "cadastra_multiplicador_associa.asp?selMultiplicadorTrue="+strMultTrue+"&selCorte="+strCorte+"&selMultiplicador="+strMult;											
				}
			}				
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadMultiplicador">		  
			 
			<!--<input type="hidden" value="<%'=strAcao%>" name="parAcao"> -->	
			<input type="hidden" name="txtCurso_Selecionados">	
			<input type="hidden" name="pintTipoMult" value="<%=intTipoMultiplicador%>">
			<input type="hidden" name="pstrTipoMult" value="<%=strTipoMultiplicador%>">
			<input type="hidden" name="parAcao" value="AS">		
									
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
						<div align="center"><a href="../../indexA_grade.asp?selCorte=<%=intCdCorte%>"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr bgcolor="#F1F1F1">
				<td colspan="3" height="20">
				  <table width="625" border="0" align="center">
					<tr>
						<td width="24"><a href="javascript:Confirma();"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
					  <td width="46"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
					  <td width="21">&nbsp;</td>
					  <td width="177"></td>
						<td width="30"></td>  
						<td width="234"></td>
					    <td width="9"></td>
					  <td width="8">&nbsp;</td>
					  <td width="38"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
					
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td height="10"></td>
				</tr>
				<tr>
				  <td>
					<div align="center"><font face="Verdana" color="#330099" size="3"><b><%=strNomeAcao%> - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="695" height="189">			  	
			  					
				<tr>
			  	  <td height="34"></td>
			  	  <td colspan="2"><font face="Verdana" color="#330099" size="2"><b>&rsaquo;&rsaquo;&nbsp;Dados do Multiplicador Extra</b></font></td>
		  	    </tr>
												
				 <tr>
					 <td height="23" colspan="1"></td>
					 <td width="131" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
					 </td>
					 <td width="376" colspan="2" valign="middle">						   				
						<input type="hidden" name="selCorte" value="<%=cint(intCdCorte)%>">	
						<font face="Verdana" size="2" color="#330099"><%=Ucase(strNomeCorte)%></font>						   	   
					</td>
			    </tr>
				
				<tr>
				  <td height="20"></td>
				  <td height="20" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>
				  <td height="20" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><%=strNomeDiretoria%></font></td>
				</tr>
				
				<tr> 
				  <td width="174" height="21"></td>
				  <td width="131" height="21" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Multiplicador:</b></font></td>
				  <td height="21" valign="middle" align="left" width="376">			 	  
					<input type="hidden" name="selMultiplicador" value="<%=Request("selMultiplicador")%>">
				  	<font face="Verdana" size="2" color="#330099"><%=Ucase(strNomeMultiplicador)%></font>				  
				  </td>
				</tr> 		
				
				<tr>
					<td width="174" height="20"></td>
					<td valign="top"><font face="Verdana" size="2" color="#330099"><b>Curso:</b></font></td>
				  <td width="376" height="20" valign="top"><font face="Verdana" size="2" color="#330099"><%=strTodosCursos%></font></td>
				</tr>		
				
				<tr>
					<td width="174" height="35"></td>
					<td colspan="2"><hr></td>
				</tr>		
		  </table>			
			<%
			'********************  INFORMAÇOES DO MULTIPLICADOR SELECIONADO *****
			if trim(Request("selMultiplicadorTrue")) <> "" then
					
				if trim(Request("selMultiplicadorTrue")) <> "" and trim(Request("selMultiplicadorTrue")) <> "0" then
					strVetMultTrue = split(trim(Request("selMultiplicadorTrue")),"|")
					intCdMultiplicadorTrue = strVetMultTrue(0)
					intTipoMultiplicadorTrue = strVetMultTrue(1)
				end if			
				
				'Response.write "intCdMultiplicador - " & intCdMultiplicador & "<br>"
				'Response.write "intCdMultiplicador - " & trim(Request("selMultiplicador")) & "<br><br>"
				'Response.write "intCdMultiplicadorTrue - " & intCdMultiplicadorTrue & "<br>"
				'Response.write "intCdMultiplicadorTrue - " & trim(Request("selMultiplicadorTrue")) & "<br>"
				
				'**************** MULTIPLICADOR SELECIONADO ****************	
				strSQLAltMultiplicadorTrue = ""
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & "SELECT CORT_CD_CORTE, MULT_NR_CD_ID_MULT, MULT_TX_NOME_MULTIPLICADOR, ORME_CD_ORG_MENOR, "
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & "MULT_TX_TIPO_MULTIPLICADOR, MULT_TX_RESTRICAO_VIAGEM, MULT_NR_TIPO_MULTIPLICADOR "
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & "FROM GRADE_MULTIPLICADOR " 
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicadorTrue
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & " AND CORT_CD_CORTE = " & intCdCorte 
				strSQLAltMultiplicadorTrue = strSQLAltMultiplicadorTrue & " AND MULT_NR_TIPO_MULTIPLICADOR = " & intTipoMultiplicadorTrue
				'Response.write strSQLAltMultiplicadorTrue & "<br><br>"
				'Response.end
				
				Set rdsAltMultiplicadorTrue = db_banco.Execute(strSQLAltMultiplicadorTrue)			
				
				if not rdsAltMultiplicadorTrue.EOF then			
				
					strMostraInformMultTrue 	= true 
					strNomeMultiplicadorTrue	= trim(rdsAltMultiplicadorTrue("MULT_TX_NOME_MULTIPLICADOR"))		
					intCdDiretoriaTrue		 	= rdsAltMultiplicadorTrue("ORME_CD_ORG_MENOR")					
					
					'************ DIRETORIA ****************
					strSQLDiretoriaTrue =  ""
					strSQLDiretoriaTrue = strSQLDiretoriaTrue & "SELECT DIRE_TX_DESC_DIRETORIA "
					strSQLDiretoriaTrue = strSQLDiretoriaTrue & "FROM GRADE_DIRETORIA "
					strSQLDiretoriaTrue = strSQLDiretoriaTrue & "WHERE ORLO_CD_ORG_LOT = " & intCdDiretoriaTrue
					'Response.WRITE  strSQLDiretoriaTrue & "<br><br>"
					'Response.END
					
					set rdsDiretoriaTrue = db_banco.execute(strSQLDiretoriaTrue)
					
					if not rdsDiretoriaTrue.eof then
						strNomeDiretoriaTrue = rdsDiretoriaTrue("DIRE_TX_DESC_DIRETORIA")
					else
						strNomeDiretoriaTrue = "N/A"
					end if
				
					rdsDiretoriaTrue.close
					set rdsDiretoriaTrue = nothing	
				
					'*** MULTIPLICADOR X CURSO - PARA O MULTIPLICADOR SELECIONADO *****				
					strNomeMultiplicadorCursoTrue = ""
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & "SELECT DISTINCT MULT_CURSO.CURS_CD_CURSO "
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & "AND MULT_CURSO.MULT_NR_CD_ID_MULT = " & intCdMultiplicadorTrue
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte 
					strNomeMultiplicadorCursoTrue = strNomeMultiplicadorCursoTrue & " AND MULT.CORT_CD_CORTE = " & intCdCorte 
					'Response.write strNomeMultiplicadorCursoTrue
					'Response.end
				
					Set rdsAltMultiplicadorTrueCurso = db_banco.Execute(strNomeMultiplicadorCursoTrue)	
												
					if not rdsAltMultiplicadorTrueCurso.EOF then					
						
						strCountCursoTrue = 0			
						strTodosCursosTrue = ""	
						'strTodosCursosQuery = ""	
						
						do while not rdsAltMultiplicadorTrueCurso.eof					
							if strCountCursoTrue = 0 then
								strTodosCursosTrue = rdsAltMultiplicadorTrueCurso("CURS_CD_CURSO")
								'strTodosCursosQuery = "'" & rdsAltMultiplicadorTrueCurso("CURS_CD_CURSO") & "'"
							else
								strTodosCursosTrue = strTodosCursosTrue & " - " & rdsAltMultiplicadorTrueCurso("CURS_CD_CURSO")
								'strTodosCursosQuery = strTodosCursosQuery & ",'" & rdsAltMultiplicadorTrueCurso("CURS_CD_CURSO") & "'"
							end if			
							
							strCountCursoTrue = strCountCursoTrue + 1
							rdsAltMultiplicadorTrueCurso.movenext
						loop		
					else				
						strTodosCursosTrue = ""		
						'strTodosCursosQuery = ""				
					end if	
				end if
			end if
			%>
				
				
			<!--- MULTIPLICADORES QUE POSSUEM O(S) CURSO(S) IGUAL AO MULTIPLICADOR EXTRA --->								 	
			<table width="694">  					
				<tr>
				  <td width="173" height="34"></td>
				  <td width="509" colspan="2">
				  	<font face="Verdana" color="#330099" size="2"><b>Multiplicador:</b></font>&nbsp;
					<select name="selMultiplicadorTrue" size="1" tabindex="1" onchange="javascript:submet_pagina(this.value,'<%=intCdCorte%>');">
						<option value="0">-- Escolha um Multiplicador ---</option>
						<%
						 if strMostraComboMultTrue = true then
							 do while not rsMultiplicador.eof
								  if trim(Request("selMultiplicadorTrue")) = trim(rsMultiplicador("MULT_NR_CD_ID_MULT") & "|" & rsMultiplicador("MULT_NR_TIPO_MULTIPLICADOR")) then
								  %>
								  <option selected value="<%=rsMultiplicador("MULT_NR_CD_ID_MULT") & "|" & rsMultiplicador("MULT_NR_TIPO_MULTIPLICADOR")%>"><%=rsMultiplicador("MULT_TX_NOME_MULTIPLICADOR")%></option>
								  <%
								  else
								   %>
								  <option value="<%=rsMultiplicador("MULT_NR_CD_ID_MULT") & "|" & rsMultiplicador("MULT_NR_TIPO_MULTIPLICADOR")%>"><%=rsMultiplicador("MULT_TX_NOME_MULTIPLICADOR")%></option>
								  <%
								  end if
								  rsMultiplicador.MoveNext
							 Loop
						end if
						%>
					</select>	
					<%
					if strMostraComboMultTrue = true then
						rsMultiplicador.close
						set rsMultiplicador = nothing
					end if
					%>
				  </td>
			    </tr>
		  </table>	
				<%
					if strMostraInformMultTrue = true then
				%>
					<table>	
					<tr>
					  <td height="34"></td>
					  <td colspan="2"><font face="Verdana" color="#330099" size="2"><b>&rsaquo;&rsaquo;&nbsp;Dados do Multiplicador</b></font></td>
					</tr>
													
					 <tr>
						 <td height="32" colspan="1"></td>
						 <td width="131" valign="middle">						
						   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
						 </td>
						 <td width="376" colspan="2" valign="middle">						   				
							<!--<input type="hidden" name="selCorte" value="<%'=cint(intCdCorte)%>">-->
							<font face="Verdana" size="2" color="#330099"><%=Ucase(strNomeCorte)%></font>						   	   
						</td>
					</tr>
					
					<tr>
					  <td height="26"></td>
					  <td height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>
					  <td height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><%=strNomeDiretoriaTrue%></font></td>
					</tr>
									
					<tr>
						<td width="174" height="25"></td>
						<td valign="top"><font face="Verdana" size="2" color="#330099"><b>Curso:</b></font></td>
					  <td width="376" height="25" valign="top"><font face="Verdana" size="2" color="#330099"><%=strTodosCursosTrue%></font></td>
					</tr>					
				
					<tr>
						<td width="174" height="21"></td>
						<td colspan="2"></td>
					</tr>		
			  </table>		 		  
	        <%
				end if
				%>  
		</form>
	</body>
	<%		
	db_banco.close
	set db_banco = nothing
	%>
</html>

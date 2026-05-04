<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
'"Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
db_banco.CursorLocation = 3

if trim(Request("selCorte")) <> "0" and trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))
end if 

strUnidade 				= request("selUnidade")
strDiretoria 			= request("selDiretoria")
strCurso 				= request("selCurso")

strDescentralizado 		= request("rdDescentralizado")
strEaD 					= request("rdEad")
strInLoco 				= request("rdInLoco")

strNumRel				= request("pNumRel")
strTituloRel			= request("pTituloRel")

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

strSQLDemanda = ""
strSQLDemanda = strSQLDemanda & "SELECT UNIDADE.UNID_CD_UNIDADE, "
strSQLDemanda = strSQLDemanda & "UNIDADE.UNID_TX_DESC_UNIDADE, "
strSQLDemanda = strSQLDemanda & "DEMANDA.CURS_CD_CURSO, "
strSQLDemanda = strSQLDemanda & "SUM(DEMANDA.DEMA_NR_TOTAL) AS TOTAL_DEMANDA, "
'strSQLDemanda = strSQLDemanda & "DIRETORIA.ORLO_CD_ORG_LOT, "
'strSQLDemanda = strSQLDemanda & "UNIDADE.ORLO_CD_ORG_LOT_DIR, "
strSQLDemanda = strSQLDemanda & "CT.CTRO_CD_CENTRO_TREINAMENTO, "
strSQLDemanda = strSQLDemanda & "DIRETORIA.DIRE_TX_DESC_DIRETORIA, "
strSQLDemanda = strSQLDemanda & "CT.CTRO_TX_NOME_CENTRO_TREINAMENTO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_NUM_CARGA_CURSO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_METODO_CURSO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_CENTRALIZADO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_IN_LOCO, "	
strSQLDemanda = strSQLDemanda & "CURSO.CURS_DT_FIM_MATERIAL_DIDATICO "
strSQLDemanda = strSQLDemanda & "FROM GRADE_UNIDADE_ORGAO_MENOR AS ORGAO_MENOR "
strSQLDemanda = strSQLDemanda & "INNER JOIN GRADE_DEMANDA_ORIGINAL_SEM AS DEMANDA ON ORGAO_MENOR.ORME_CD_ORG_MENOR = DEMANDA.ORME_CD_ORG_MENOR "
strSQLDemanda = strSQLDemanda & "INNER JOIN GRADE_UNIDADE AS UNIDADE ON ORGAO_MENOR.UNID_CD_UNIDADE = UNIDADE.UNID_CD_UNIDADE "
strSQLDemanda = strSQLDemanda & "INNER JOIN GRADE_DIRETORIA AS DIRETORIA ON UNIDADE.ORLO_CD_ORG_LOT_DIR = DIRETORIA.ORLO_CD_ORG_LOT "
strSQLDemanda = strSQLDemanda & "INNER JOIN GRADE_CENTRO_TREINAMENTO AS CT ON UNIDADE.CTRO_CD_CENTRO_TREINAMENTO = CT.CTRO_CD_CENTRO_TREINAMENTO "
strSQLDemanda = strSQLDemanda & "INNER JOIN GRADE_CURSO AS CURSO ON DEMANDA.CURS_CD_CURSO = CURSO.CURS_CD_CURSO "
strSQLDemanda = strSQLDemanda & "WHERE DEMANDA.CORT_CD_CORTE = " & Session("Corte")
strSQLDemanda = strSQLDemanda & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte")
strSQLDemanda = strSQLDemanda & " AND CT.CORT_CD_CORTE = " & Session("Corte")
strSQLDemanda = strSQLDemanda & " AND CURSO.CORT_CD_CORTE = " & Session("Corte")

if strDiretoria <> "0" then
	strSQLDemanda = strSQLDemanda & " AND UNIDADE.ORLO_CD_ORG_LOT_DIR = " & strDiretoria 
end if

if strUnidade <> "0" then
	strSQLDemanda = strSQLDemanda & " AND UNIDADE.UNID_CD_UNIDADE = '" & strUnidade & "'"
end if 

if strCurso <> "0" then
	strSQLDemanda = strSQLDemanda & " AND DEMANDA.CURS_CD_CURSO = '" & strCurso & "'"
end if 

if strEaD = "0" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_METODO_CURSO = 'PRESENCIAL'"
elseif strEaD = "1" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_METODO_CURSO = 'À DISTÂNCIA'"
end if

if strDescentralizado = "0" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_CENTRALIZADO = 'CENTRALIZADO'"
elseif strDescentralizado = "1" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_CENTRALIZADO = 'DESCENTRALIZADO'"
end if

if strInLoco = "0" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_IN_LOCO = 'S'"
elseif strInLoco = "1" then
	strSQLDemanda = strSQLDemanda & " AND CURSO.CURS_TX_IN_LOCO = 'N'"
end if

strSQLDemanda = strSQLDemanda & " GROUP BY UNIDADE.UNID_CD_UNIDADE, "
strSQLDemanda = strSQLDemanda & "UNIDADE.UNID_TX_DESC_UNIDADE, "
strSQLDemanda = strSQLDemanda & "DEMANDA.CURS_CD_CURSO, "
'strSQLDemanda = strSQLDemanda & "DIRETORIA.ORLO_CD_ORG_LOT, "
strSQLDemanda = strSQLDemanda & "CT.CTRO_CD_CENTRO_TREINAMENTO, "
strSQLDemanda = strSQLDemanda & "DIRETORIA.DIRE_TX_DESC_DIRETORIA, "
strSQLDemanda = strSQLDemanda & "CT.CTRO_TX_NOME_CENTRO_TREINAMENTO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_CD_CURSO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_NUM_CARGA_CURSO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_METODO_CURSO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_CENTRALIZADO, "                     
strSQLDemanda = strSQLDemanda & "CURSO.CURS_TX_IN_LOCO, "
strSQLDemanda = strSQLDemanda & "CURSO.CURS_DT_FIM_MATERIAL_DIDATICO "
strSQLDemanda = strSQLDemanda & " ORDER BY DIRETORIA.DIRE_TX_DESC_DIRETORIA "
'Response.write strSQLDemanda
'Response.end		
		
set rstDemanda = db_banco.execute(strSQLDemanda)				
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<style type="text/css">
			<!--
			.style2 
			{
				font-family: Verdana, Arial, Helvetica, sans-serif;
				font-weight: bold;
				color: #000066;
			}		
			-->
		</style>
	</head>
	
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">		
		<%		
		if request("excel") <> 1 then
		%>		
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
					<td width="27"></td>  
					<td width="50"></td>
				  <td width="28"></td>
				  <td width="26">&nbsp;</td>
				  <td width="159"></td>
				</tr>
			  </table>
			</td>
		  </tr>
		</table>	
		
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="720"></td>	
					<td width="50">
						<div align="center">	
							 <a href="gera_consulta_demanda.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;rdDescentralizado=<%=strDescentralizado%>&amp;rdEad=<%=strEaD%>&amp;rdInLoco=<%=strInLoco%>&amp;selCurso=<%=strCurso%>&amp;pTituloRel=<%=strTituloRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
						</div>
					</td>						
				</tr>
			</table>
		<%end if%>		
		
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3"><b>Relatório de <%=strTituloRel%> - Grade de Treinamento</b></font></div>
			  </td>
			</tr>
				
			<tr>
			  <td>&nbsp;</td>
			</tr>
    </table>		
				
		<table width="99%" border="0" cellpadding="2" cellspacing="2">
		  <tr bgcolor="#D4D0C8">
		    <td width="9" bgcolor="#ffffff"></td>
		    <td width="85" align="center">
					<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Diretoria</b></font>
			</td>
			<%'if strUnidade = "0" then%>
				<td width="78" align="center">
					<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Unidade</b></font>
				</td>
			<%'end if%>									
		    <td width="99" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Centro de Treinamento</b></font>
			</td>		  
		     <td width="66" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Cód Curso</b></font>
			</td>
			 <td width="86" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Data Fim<br>Material Didático</b></font>
			</td>
			
			 <td width="49" align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Qtde de Usu&aacute;rios</b></font></td>
			 
			<td width="40" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Carga Horária</b></font>
			</td>
			
			<td width="151" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Pré-Requisito/Método</b></font></td>
				    
			<td width="122" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Tipo do Curso</b></font>
			</td>
			
		    <td width="92" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Método do Curso</b></font>
			</td>
			
		    <td width="48" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>In Loco</b></font>
			</td>				
		  </tr>		
			<%
			'*** INICIALIZAÇÕES ***	
			strCor = "#FFFFFF"
			intTotal = 0			
			'strNomeAtual = ""
			strNomeCurso = ""
			'intQtdeUsuario = 0
			intTotalQtdeUsuario = 0
			
			if not rstDemanda.eof then									
				do until rstDemanda.eof = true
														
					if strCor = "#FFFFFF" then
						strCor = "#EAEAEA"
					else
						strCor = "#FFFFFF"
					end if
					
					'if strNomeAtual = "" then
						strNomeCurso = trim(rstDemanda("CURS_CD_CURSO"))
						'strNomeAtual = strNomeCurso
						
						'intQtdeUsuario = rstDemanda("TOTAL_DEMANDA")
						intTotalQtdeUsuario = intTotalQtdeUsuario + rstDemanda("TOTAL_DEMANDA")						
					'else
					'	if strNomeAtual = trim(rstDemanda("CURS_CD_CURSO")) then
						'	'strNomeCurso = ""
							'intQtdeUsuario = ""
						'else
							'strNomeCurso = trim(rstDemanda("CURS_CD_CURSO"))
							'strNomeAtual = strNomeCurso
							
							'intQtdeUsuario = rstDemanda("TOTAL_DEMANDA")
							'intTotalQtdeUsuario = intTotalQtdeUsuario + rstDemanda("TOTAL_DEMANDA")						
						'end if						
					'end if				
										
					if rstDemanda("CURS_TX_CENTRALIZADO") <> "" then
						strTxtTipo = rstDemanda("CURS_TX_CENTRALIZADO")
					else
						strTxtTipo = "N/A"
					end if	
													
					if rstDemanda("CURS_TX_METODO_CURSO") <> "" then
						strTxtCentralizada = ucase(rstDemanda("CURS_TX_METODO_CURSO"))
					else
						strTxtCentralizada = "N/A"
					end if												
											
					if rstDemanda("CURS_TX_IN_LOCO") = "S" then
						strTxtInLoco = "SIM"
					elseif rstDemanda("CURS_TX_IN_LOCO") = "N" then
						strTxtInLoco = "NÃO"
					else
						strTxtInLoco = "N/A"
					end if																				
					%>					
					<tr bgcolor="<%=strCor%>">
						<td width="9" bgcolor="#ffffff"></td>
						<td width="85" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=Ucase(rstDemanda("DIRE_TX_DESC_DIRETORIA"))%></font>
						</td>
					
						<td width="78" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=Ucase(rstDemanda("UNID_TX_DESC_UNIDADE"))%></font>
						</td>
												
						<td width="99" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=Ucase(rstDemanda("CTRO_TX_NOME_CENTRO_TREINAMENTO"))%></font>
						</td>
						
						<td width="66" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strNomeCurso%></font>
						</td>	
						
						<%
						if trim(rstDemanda("CURS_DT_FIM_MATERIAL_DIDATICO")) <> "" then
							strDTFimMatDidatico = MontaDataHora(trim(rstDemanda("CURS_DT_FIM_MATERIAL_DIDATICO")),2)
						else
							strDTFimMatDidatico = "N/A"
						end if
						%>					
						<td width="86" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strDTFimMatDidatico%></font>
						</td>
						
						<td width="49" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstDemanda("TOTAL_DEMANDA")%></font>
						</td>
												
						<td width="40" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstDemanda("CURS_NUM_CARGA_CURSO")%></font>
						</td>
						
						<%
						strCursoPre = ""
						strCursoPre = strCursoPre & "SELECT CURS_PRE_REQUISITO "
						strCursoPre = strCursoPre & "FROM GRADE_CURSO_PRE_REQUISITO "
						strCursoPre = strCursoPre & "WHERE CURS_CD_CURSO = '" & strNomeCurso & "'"
						strCursoPre = strCursoPre & " AND CORT_CD_CORTE = " & Session("Corte")
						strCursoPre = strCursoPre & " ORDER BY CURS_CD_CURSO"
						'Response.write strCursoPre & "<br><br>"
						'Response.end
						
						set rsCursoPre = db_banco.execute(strCursoPre)
						
						strCursoPreRequisito = ""
						
						if not rsCursoPre.eof then
							
							do while not rsCursoPre.eof
							
								strCursoMetodoPre = ""
								strCursoMetodoPre = strCursoMetodoPre & "SELECT CURS_TX_METODO_CURSO "
								strCursoMetodoPre = strCursoMetodoPre & "FROM GRADE_CURSO "
								strCursoMetodoPre = strCursoMetodoPre & "WHERE CURS_CD_CURSO = '" & trim(rsCursoPre("CURS_PRE_REQUISITO")) & "'"
								strCursoMetodoPre = strCursoMetodoPre & " AND CORT_CD_CORTE = " & Session("Corte")
								'Response.write strCursoMetodoPre & "<br><br>"
								'Response.end							
															
								set rsCursoMetodoPre = db_banco.execute(strCursoMetodoPre)
																
								if not rsCursoMetodoPre.eof then							
									if strCursoPreRequisito = "" then
										strCursoPreRequisito = rsCursoPre("CURS_PRE_REQUISITO") & " - " & rsCursoMetodoPre("CURS_TX_METODO_CURSO")
									else
										strCursoPreRequisito = strCursoPreRequisito & "<br>" & rsCursoPre("CURS_PRE_REQUISITO") & " - " & rsCursoMetodoPre("CURS_TX_METODO_CURSO")
									end if
								else
									if strCursoPreRequisito = "" then
										strCursoPreRequisito = rsCursoPre("CURS_PRE_REQUISITO") 
									else
										strCursoPreRequisito = strCursoPreRequisito & "<br>" & rsCursoPre("CURS_PRE_REQUISITO")
									end if
								end if
								
								rsCursoPre.movenext
								
								rsCursoMetodoPre.close
								set rsCursoMetodoPre = nothing
							loop							
						else
							strCursoPreRequisito = "N/A"
						end if
						
						rsCursoPre.close
						set rsCursoPre = nothing						
						%>						
						<td width="151" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strCursoPreRequisito%></font>
						</td>						
						
						<td width="122" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strTxtTipo%></font>
						</td>
						
						<td width="92" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strTxtCentralizada%></font>
						</td>
						
						<td width="48" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strTxtInLoco%></font>
						</td>								
					</tr>					
					<%
					intTotal = intTotal + 1					
					rstDemanda.movenext					
				loop
				%>
				<tr><td height="20"></td></tr>
				<tr>		
					<td width="9" bgcolor="#ffffff"></td>		
					<td colspan="11">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total de Usuários:</b></font>&nbsp;&nbsp;  
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><%=intTotalQtdeUsuario%></b></font>				  
					</td>					
				</tr>				
				<tr><td height="10"></td></tr>
				<%
			else
			%>
				<tr>
					<td width="9" bgcolor="#ffffff"></td>
					<td colspan="11">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não existe registros para esta consulta.</font>
					</td>
				</tr>					
			<%
			end if
			
			rstDemanda.close
			set rstDemanda = nothing			
			%>
	</table>		
	</body>	
	<%	
	public function MontaDataHora(strData,intDataTime)
	
		'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
		'*** intDataTime = 1 (DATA E HORA)
		'*** intDataTime = 2 (DATA)
		'*** intDataTime = 3 (HORA)
		'*** intDataTime = 4 (FORMATO DE BANCO)
		'*** intDataTime = 5 (FORMATO DE BANCO - DIA E MÊS)
	
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
		elseif cint(intDataTime) = 5 then
			MontaDataHora = strDia & "/" & strMes
		end if
	end function


	db_banco.close
	set db_banco = nothing
	%>		
</html>

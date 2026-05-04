
<!--#include file="../../asp/protege/protege.asp" -->
<%
if request("excel") = 1 then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

strMega 	= Request("selMegaProcesso")
strOnda 	= Request("selOnda")
strStatus 	= Request("hdStatus")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql = ""
ssql = ssql + " SELECT distinct MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
ssql = ssql + " CURSO.CURS_CD_CURSO, "
ssql = ssql + " CURSO.ONDA_CD_ONDA, "
ssql = ssql + " 	(SELECT DISTINCT ABRANGENCIA_CURSO.ONDA_TX_DESC_ONDA "
ssql = ssql + " 	FROM ABRANGENCIA_CURSO "
ssql = ssql + " 	WHERE ABRANGENCIA_CURSO.ONDA_CD_ONDA = CURSO.ONDA_CD_ONDA) AS TX_ONDA, "
ssql = ssql + " CURSO.CURS_TX_NOME_CURSO, "
ssql = ssql + " CURSO.CURS_NUM_CARGA_CURSO, "
ssql = ssql + " CURSO.CURS_TX_METODO_CURSO "
'ssql = ssql + " FROM dbo.CURSO, dbo.CURSO_FUNCAO, dbo.CURSO_PRE_REQUISITO , dbo.MEGA_PROCESSO "
ssql = ssql + " FROM dbo.CURSO, dbo.CURSO_FUNCAO, dbo.MEGA_PROCESSO "
ssql = ssql + " WHERE CURSO.CURS_CD_CURSO = dbo.CURSO_FUNCAO.CURS_CD_CURSO "
'ssql = ssql + " AND dbo.CURSO.CURS_CD_CURSO = dbo.CURSO_PRE_REQUISITO.CURS_CD_CURSO "
ssql = ssql + " AND dbo.CURSO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "

if strOnda <> "0" then
	ssql = ssql + " AND dbo.CURSO.ONDA_CD_ONDA = " & strOnda
end if

if strMega <> "0" then
	ssql = ssql + " AND dbo.CURSO.MEPR_CD_MEGA_PROCESSO = " & strMega	
end if

if strStatus = "1" then '*** ATIVOS	
	ssql = ssql + " AND CURS_TX_STATUS_CURSO = '1'"	
elseif strStatus = "2" then	'*** INATIVOS	
	ssql = ssql + " AND CURS_TX_STATUS_CURSO = '0'"	
end if

ssql = ssql + " ORDER BY dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, dbo.CURSO.CURS_CD_CURSO "
'Response.write ssql

set rstRegistros = DB.execute(SSQL)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>	
		<style type="text/css">			
			.style7 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10; color: #FFFFFF; }
			.style9 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10; }			
		</style>
	</head>
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="post" action="" name="frm1">
		
			<%if request("excel") <> 1 then%>
		
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
						<div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr bgcolor="#00FF99">
				<td colspan="3" height="20">
				  <table width="625" border="0" align="center">
					<tr>
						<td width="26"></td>
					  	<td width="50"></td>
					  	<td width="26"></td>
					  	<td width="195"></td>
						<td width="27"></td>						
						<td width="50"><a href="rel_catalogo_curso.asp?excel=1&selMegaProcesso=<%=strMega%>&selOnda=<%=strOnda%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
					  	<td width="28"></td>
					  	<td width="26">&nbsp;</td>
					  	<td width="159"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
			
			<%end if%>
			
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td>
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center">
					  <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório de Católogo de Cursos</font>
					  
					  <table width="983" border="0" bordercolor="#333333" cellpadding="0" cellspacing="0">	
					  		<tr>
								<td width="158" height="10"></td>
								<td width="559" height="10"></td>
								<td width="266" height="10"></td>
							</tr>
							
							  <%if strMega <> "0" then							  
								  sqlMega = ""
								  sqlMega = sqlMega & "SELECT MEPR_TX_DESC_MEGA_PROCESSO FROM MEGA_PROCESSO "
								  sqlMega = sqlMega & "WHERE MEPR_CD_MEGA_PROCESSO = " & strMega						  
								  set rstMega = db.execute(sqlMega)
								  
								  if not rstMega.eof then
									strNomeMega = rstMega("MEPR_TX_DESC_MEGA_PROCESSO")
								  end if
								  
								  rstMega.close
								  set rstMega = nothing
								  %>
								<tr>									
									<td width="158"><font face="Verdana" color="#330099" size="2"><b>Mega-Processo:&nbsp;</b></font></td>
									<td width="559"><font face="Verdana" color="#330099" size="2"><%=strNomeMega%></font></td>									
								</tr>
							<%end if%>
							
							<%if strOnda <> "0" then							
								  sqlOnda = ""							  
								  sqlOnda = sqlOnda & "SELECT ONDA_TX_DESC_ONDA "
								  sqlOnda = sqlOnda & "FROM ABRANGENCIA_CURSO "
								  sqlOnda = sqlOnda & "WHERE ONDA_CD_ONDA = " & strOnda									  
								  set rstOnda = db.execute(sqlOnda)
								  
								  if not rstOnda.eof then
									strNomeOnda = rstOnda("ONDA_TX_DESC_ONDA")
								  end if
								  
								  rstOnda.close
								  set rstOnda = nothing
								  %>	
								 <tr>									
									<td><font face="Verdana" color="#330099" size="2"><b>Abrang&ecirc;ncia:&nbsp;</b></font></td>
									<td><font face="Verdana" color="#330099" size="2"><%=strNomeOnda%></font></td>
								 </tr>
							<%end if 
							
							if strStatus = "0" then '*** TODOS	
								strNomeStatus = "ATIVOS E INATIVOS"
							elseif strStatus = "1" then '*** ATIVOS	
								strNomeStatus = "ATIVOS"
							elseif strStatus = "2" then	'*** INATIVOS	
								strNomeStatus = "INATIVOS"
							end if
							%>
							<tr>
								<td width="158"><font face="Verdana" color="#330099" size="2"><b>Status do Curso:&nbsp;</b></font></td>
								<td width="559"><font face="Verdana" color="#330099" size="2"><%=strNomeStatus%></font></td>								
							</tr>
							<tr>
								<td width="158" height="10"></td>
								<td width="559" height="10"></td>
								<td width="266" height="10"></td>
							</tr>
					  </table>	  
					  
					  <%if not rstRegistros.eof then%>
					  
					  <table width="100%" border="0">					
						
						<tr bgcolor="#31009C">
						
						<%						
						if strMega = "0" then%>						
						  <td width="168"><span class="style7">Mega-Processo</span></td>
						<%end if%> 
						
						  <td width="49"><span class="style7">Curso</span></td>
						  <td width="209"><span class="style7">Nome</span></td>
						  <td width="59"><span class="style7">Carga Hor&aacute;ria</span></td>
						  <td width="72"><span class="style7">M&eacute;todo</span></td>
						
						<%if strOnda = "0" then%>	
						  <td width="134"><span class="style7">Abrang&ecirc;ncia</span></td>
						<%end if%> 
						  <!--<td width="75"><span class="style7">Requisitos n&atilde;o R/3 </span></td>-->
						   <td width="124"><span class="style7">Pré-Requisito</span></td>
						   <td width="151" align="justify"><span class="style7">P&uacute;blico Alvo </span></td>						  
						</tr>
						<%						
						anterior=""
						atual=""
						do until rstRegistros.eof = true
							atual = rstRegistros("CURS_CD_CURSO")
							
							if COR="WHITE" then
								COR="#E4E4E4"
							else
								COR="WHITE"
							end if
							%>
							<tr>
							  <%if strMega = "0" then%>	
								<td bgcolor="<%=COR%>"><span class="style9"><%=rstRegistros("MEPR_TX_DESC_MEGA_PROCESSO")%></span></td>
							  <%end if%>
							  
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rstRegistros("CURS_CD_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rstRegistros("CURS_TX_NOME_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>" align="center" width="59"><span class="style9"><%=rstRegistros("CURS_NUM_CARGA_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rstRegistros("CURS_TX_METODO_CURSO")%></span></td>
							  
							  <%if strOnda = "0" then%>
								<td bgcolor="<%=COR%>"><span class="style9"><%=rstRegistros("TX_ONDA")%></span></td>
							  <%end if%>
							  
							  <%
							  sqlComplCurso = ""
							  sqlComplCurso = sqlComplCurso & "SELECT CURS_TX_PUBLICO_ALVO, CURS_TX_PRE_REQUISITOS FROM CURSO "
							  sqlComplCurso = sqlComplCurso & "WHERE CURS_CD_CURSO = '" & rstRegistros("CURS_CD_CURSO") & "'"
							 							  
							  set rstComplCurso = db.execute(sqlComplCurso)
								  
							  if not rstComplCurso.eof then
								strNomeComplCurso = rstComplCurso("CURS_TX_PUBLICO_ALVO")
																
								'if rstComplCurso("CURS_TX_PRE_REQUISITOS") <> "" then
								'	strPreRequisitoN_R3 = rstComplCurso("CURS_TX_PRE_REQUISITOS")
								'else
								'	strPreRequisitoN_R3 = "N/A"
								'end if								
							  end if
							  
							  rstComplCurso.close
							  set rstComplCurso = nothing 
							  
							  sqlCursoPreRequisito = ""
							  sqlCursoPreRequisito = sqlCursoPreRequisito & "SELECT CURS_PRE_REQUISITO "
							  sqlCursoPreRequisito = sqlCursoPreRequisito & "FROM CURSO_PRE_REQUISITO "
							  sqlCursoPreRequisito = sqlCursoPreRequisito & "WHERE CURS_CD_CURSO = '" & rstRegistros("CURS_CD_CURSO") & "'"
							 							  
							  set rstCursoPreRequisito = db.execute(sqlCursoPreRequisito)
							  
							  strPreRequisito = ""  
							  if not rstCursoPreRequisito.eof then						  																
								do while not rstCursoPreRequisito.eof 
									if strPreRequisito = "" then
										strPreRequisito = rstCursoPreRequisito("CURS_PRE_REQUISITO")
									else
										strPreRequisito = strPreRequisito & " " & rstCursoPreRequisito("CURS_PRE_REQUISITO")
									end if		
									rstCursoPreRequisito.movenext
								loop					
							  else								
								strPreRequisito = "N/A"														
							  end if
							  
							  rstCursoPreRequisito.close
							  set rstCursoPreRequisito = nothing 
							  
							  %>			
							  <!--<td bgcolor="<%'=COR%>"><span class="style9"><%'=strPreRequisitoN_R3%></span></td>-->		  
							  <td bgcolor="<%=COR%>"><span class="style9"><%=strPreRequisito%></span></td>	
							  <td bgcolor="<%=COR%>"><span class="style9"><%=StrNomeComplCurso%></span></td>						
							</tr>
							<%
							if anterior <> atual then
								tem = tem + 1
							end if
							anterior = rstRegistros("CURS_CD_CURSO")
							rstRegistros.MOVENEXT
						LOOP
						%>
					  </table>					  			  
					  <p align="left" class="style9">Total de Cursos Dispon&iacute;veis:&nbsp;<strong><%=tem%></strong></p>
					  <p align="left"></p> 					  
					</div>
				  </td>
				</tr>
			  </table>
			  		
			  <%else%>      
					<p align="center"><font face="Verdana" color="#330099" size="2">Não existe resultado para esta consulta.</font></p>
				   	<p align="left"></p>  
			  <%end if%>
		</form>	
	</body>
</html>

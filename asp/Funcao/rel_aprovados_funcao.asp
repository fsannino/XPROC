<%
set conn_Cogest = Server.CreateObject("ADODB.Connection")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.cursorlocation = 3

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
else
	response.Buffer = false
	server.ScriptTimeout = 3600
end if

strSQL = ""
strSQL = strSQL & " SELECT O_AGLU.AGLU_SG_AGLUTINADO, "
strSQL = strSQL & "	O_MAIOR.ORLO_SG_ORG_LOT, "
strSQL = strSQL & "	O_MENOR.ORME_SG_ORG_MENOR, "
strSQL = strSQL & " F_USUA.USMA_CD_USUARIO, "
strSQL = strSQL & "	U_MAP.USMA_TX_NOME_USUARIO, "
strSQL = strSQL & " MP.MEPR_TX_DESC_MEGA_PROCESSO, "
strSQL = strSQL & " F_USUA.FUNE_CD_FUNCAO_NEGOCIO, "
strSQL = strSQL & " F_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO, "
strSQL = strSQL & " F_USUA.FUUS_IN_PRIORITARIO, "
strSQL = strSQL & " F_USUA.FUUS_IN_VALIDADO, "
strSQL = strSQL & " C_FUNCAO.CURS_CD_CURSO, "
strSQL = strSQL & " U_APROV.USAP_TX_APROVEITAMENTO, "
strSQL = strSQL & " M_PERFIL.MIPE_TX_DESC_MICRO_PERFIL, "
strSQL = strSQL & " FUNC_U_PERFIL.FUUP_IN_VALIDADO "
strSQL = strSQL & " FROM ORGAO_MAIOR AS O_MAIOR, "
strSQL = strSQL & " ORGAO_MENOR AS O_MENOR, "
strSQL = strSQL & " USUARIO_MAPEAMENTO AS U_MAP, "
strSQL = strSQL & " FUNCAO_NEGOCIO AS F_NEG, "
strSQL = strSQL & " FUNCAO_USUARIO AS F_USUA, "
strSQL = strSQL & " MEGA_PROCESSO AS MP, "
strSQL = strSQL & " ORGAO_AGLUTINADOR AS O_AGLU, "
strSQL = strSQL & " CURSO_FUNCAO AS C_FUNCAO, "
strSQL = strSQL & " USUARIO_APROVADO AS U_APROV, "
strSQL = strSQL & " FUNCAO_USUARIO_PERFIL AS FUNC_U_PERFIL,"
strSQL = strSQL & " MICRO_PERFIL_R3 AS M_PERFIL "	
strSQL = strSQL & " WHERE O_MENOR.ORME_CD_ORG_MENOR 		= U_MAP.ORME_CD_ORG_MENOR "
strSQL = strSQL & " AND U_MAP.USMA_CD_USUARIO 			= F_USUA.USMA_CD_USUARIO "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		= F_NEG.FUNE_CD_FUNCAO_NEGOCIO "
strSQL = strSQL & " AND O_MAIOR.ORLO_CD_ORG_LOT 			= O_MENOR.ORLO_CD_ORG_LOT "
strSQL = strSQL & " AND F_NEG.MEPR_CD_MEGA_PROCESSO 		= MP.MEPR_CD_MEGA_PROCESSO "
strSQL = strSQL & " AND O_MAIOR.AGLU_CD_AGLUTINADO 		= O_AGLU.AGLU_CD_AGLUTINADO "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		= C_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO "
strSQL = strSQL & " AND C_FUNCAO.CURS_CD_CURSO 			= U_APROV.CURS_CD_CURSO "
strSQL = strSQL & " AND F_USUA.USMA_CD_USUARIO 			= U_APROV.USAP_CD_USUARIO "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		= FUNC_U_PERFIL.FUNE_CD_FUNCAO_NEGOCIO "
strSQL = strSQL & " AND F_USUA.USMA_CD_USUARIO 			= FUNC_U_PERFIL.USMA_CD_USUARIO "
strSQL = strSQL & " AND FUNC_U_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL 	= M_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
strSQL = strSQL & " AND FUNC_U_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL 	= M_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.01' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.02' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.03'"
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.04' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.05' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.07' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.08' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.11' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.17' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.21' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.22' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.24' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.025' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.027' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.028' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.029' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.030' "
strSQL = strSQL & " AND F_USUA.FUNE_CD_FUNCAO_NEGOCIO 		<> 'HR.031' "
strSQL = strSQL & " AND F_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO 	NOT LIKE '%(ANT%)'"
strSQL = strSQL & " AND F_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO 	NOT LIKE '%(TRANSIT%)' "
strSQL = strSQL & " AND O_MAIOR.ORLO_CD_STATUS 			= 'A' "
strSQL = strSQL & " AND O_MENOR.ORME_CD_STATUS 			= 'A' "
'strSQL = strSQL & " AND O_MAIOR.ORLO_SG_ORG_LOT			= '' "
'strSQL = strSQL & " AND  F_USUA.USMA_CD_USUARIO = 'DLS1' "
strSQL = strSQL & " GROUP BY O_AGLU.AGLU_SG_AGLUTINADO, O_MAIOR.ORLO_SG_ORG_LOT, "
strSQL = strSQL & " O_MENOR.ORME_SG_ORG_MENOR, F_USUA.USMA_CD_USUARIO, U_MAP.USMA_TX_NOME_USUARIO, "
strSQL = strSQL & " MP.MEPR_TX_DESC_MEGA_PROCESSO, F_USUA.FUNE_CD_FUNCAO_NEGOCIO, "
strSQL = strSQL & " F_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO, F_USUA.FUUS_IN_PRIORITARIO, "
strSQL = strSQL & " F_USUA.FUUS_IN_VALIDADO, C_FUNCAO.CURS_CD_CURSO, U_APROV.USAP_TX_APROVEITAMENTO," 
strSQL = strSQL & " M_PERFIL.MIPE_TX_DESC_MICRO_PERFIL, FUNC_U_PERFIL.FUUP_IN_VALIDADO, "
strSQL = strSQL & " O_MAIOR.ORLO_CD_STATUS, O_MENOR.ORME_CD_STATUS"
strSQL = strSQL & " ORDER BY O_AGLU.AGLU_SG_AGLUTINADO, O_MAIOR.ORLO_SG_ORG_LOT, O_MENOR.ORME_SG_ORG_MENOR, " 
strSQL = strSQL & " F_USUA.USMA_CD_USUARIO, MP.MEPR_TX_DESC_MEGA_PROCESSO, "
strSQL = strSQL & " F_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
'Response.write strSQL
'Response.end

Set rds_Usu_Curso = conn_Cogest.Execute(strSQL)
int_NumReg_Usu_Curso = rds_Usu_Curso.RecordCount
'Response.write int_NumReg_Usu_Curso & "<br>"
'Response.end
%>
<html>
	<head>		
		<STYLE type=text/css>
			BODY {
				SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
		</STYLE>
		<style>
			a {text-decoration:none;}
			a:hover {text-decoration:underline;}
		</style>
		<title>SINERGIA # XPROC # Processos de Negócio</title>	
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
	</head>
	<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">		
		
			<%if request("excel") <> 1 then%>					
					 <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
						<tr>
						  <td width="160" height="20" colspan="2">&nbsp;</td>
						  <td width="346" height="60" colspan="3">&nbsp;</td>
						  <td width="330" valign="top" colspan="2"> 
							<table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
							<tr> 
							  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
								<div align="center">
								  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a></div>
							  </td>
							  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
								<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
							  </td>
							  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
								<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
							  </td>
							</tr>
							<tr> 
							  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
								<div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
							  </td>
							  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
								<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
							  </td>
							  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
								  <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
							  </td>
							</tr>
						  </table>
						</td>
					  </tr>
					  <tr bgcolor="#00FF99"> 
						<td height="20" width="155">&nbsp;</td>
						<td colspan="2" height="20" width="31">&nbsp;</td>
						<td height="20" width="244">&nbsp;</td>
						<td colspan="2" height="20" width="112"><a href="aprovados_funcao.asp?excel=1" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
						<td height="20" width="290">&nbsp;</td>
					  </tr>
					</table>			
			<%end if%>	
					
		<table width="100%" border="0">					
			<tr><td colspan="13" width="100%" height="10"></td></tr>
			
			<tr>
				<td colspan="13" align="center" width="100%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000080" size="2"><b>Relatório de Aprovados em Cursos por Função</b></font></td>
			</tr>
			
			<tr><td colspan="13" width="100%">
				<img src="../../imagens/carregando01.gif" name="imagem1" width="120" height="18" id="imagem1">
			</td></tr>	
		  <tr>
			<td width="12%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Orgão Aglutinador</b></font></td>
			<td width="8%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Orgão Maior</b></font></td>
			<td width="8%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Orgão Menor</b></font></td>
			<td width="5%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Chave</b></font></td>
			<td width="11%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Nome</b></font></td>
			<td width="10%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Mega-Processo</b></font></td>
			<td width="6%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Codigo</b></font></td>
			<td width="6%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Função</b></font></td>
			<td width="7%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Prioritário</b></font></td>
			<td width="7%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Validado</b></font></td>
			<td width="5%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Curso</b></font></td>
			<td width="8%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Perfil</b></font></td>
			<td width="7%" align="center" bgcolor="#000080"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1"><b>Validado Perfil</b></font></td>
		  </tr>	
		<%	
		intTotalRegistros = 0
		int_Loop_Usu_Curso = 0
		If int_NumReg_Usu_Curso > 0 Then
			'*** LOOP EM TODOS OS REGISTROS
			Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso
			
				str_Cd_Usu_Anterior = rds_Usu_Curso("USMA_CD_USUARIO")
				
				'*** LOOP EM TODOS OS REGISTROS ENQUANTO MESMO USUARIO			
				Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso Or str_Cd_Usu_Anterior <> rds_Usu_Curso("USMA_CD_USUARIO")
					str_Cd_Fun_Anterior = rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
					int_Nao_Aprovado = 0
					int_Aprovado = 0
					
					'*** LOOP EM TODOS OS REGISTROS ENQUANTO MESMO USUARIO E FUNCÃO					
					Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso Or str_Cd_Usu_Anterior <> rds_Usu_Curso("USMA_CD_USUARIO") Or str_Cd_Fun_Anterior <> rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
						
						str_SQL = " SELECT "
						str_SQL = str_SQL & " USAP_TX_APROVEITAMENTO "
						str_SQL = str_SQL & " FROM  USUARIO_APROVADO "
						str_SQL = str_SQL & " WHERE USAP_CD_USUARIO ='" & str_Cd_Usu_Anterior & "'"
						str_SQL = str_SQL & " And CURS_CD_CURSO ='" & rds_Usu_Curso("CURS_CD_CURSO") & "'"
						'Response.Write str_SQL
						'Response.end	
						
						Set rds_TabIncr = conn_Cogest.Execute(str_SQL)
						int_NumReg_TabIncr = rds_TabIncr.RecordCount
						'Response.Write int_NumReg_TabIncr
						'Response.end				
						
						int_Loop_TabIncr = 0				
						if int_NumReg_TabIncr > 0 Then
							Do Until int_NumReg_TabIncr = int_Loop_TabIncr
								If rds_TabIncr("USAP_TX_APROVEITAMENTO") = "AP" or rds_TabIncr("USAP_TX_APROVEITAMENTO") = "LM" then
									int_Aprovado = int_Aprovado + 1
								else
									int_Nao_Aprovado = int_Nao_Aprovado + 1							
								end if
								int_Loop_TabIncr = int_Loop_TabIncr + 1
								rds_TabIncr.MoveNext
							Loop
						else
							int_Nao_Aprovado = int_Nao_Aprovado + 1
						end If
						
						rds_Usu_Curso.MoveNext
						int_Loop_Usu_Curso = int_Loop_Usu_Curso + 1
						
						If rds_Usu_Curso.EOF Then
						   Exit Do
						End If
						rds_TabIncr.Close
					Loop
					'****************************
					'int_Nao_Aprovado = 0
					'****************************
					If int_Nao_Aprovado = 0 Then		
					
					  if cor="#E4E4E4" then
						cor="white"
					  else
						cor="#E4E4E4"
					  end if
						If not rds_Usu_Curso.EOF Then	
							intTotalRegistros = intTotalRegistros + 1	
							%>
							 <tr>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("AGLU_SG_AGLUTINADO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("ORLO_SG_ORG_LOT")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("ORME_SG_ORG_MENOR")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><%=rds_Usu_Curso("USMA_CD_USUARIO")%></b></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><%=rds_Usu_Curso("USMA_TX_NOME_USUARIO")%></b></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("FUUS_IN_PRIORITARIO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("FUUS_IN_VALIDADO")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><%=rds_Usu_Curso("CURS_CD_CURSO")%></b></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("MIPE_TX_DESC_MICRO_PERFIL")%></font></td>
								<td bgcolor="<%=cor%>" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><%=rds_Usu_Curso("FUUP_IN_VALIDADO")%></font></td>
							  </tr>
							<%	
						end if										
					'Else						
					   'Response.write rds_Usu_Curso("USMA_CD_USUARIO") & " - Reprovado<br>"					   
					End If
					If rds_Usu_Curso.EOF Then
					   Exit Do
					End If
				Loop
				If rds_Usu_Curso.EOF Then
				   Exit Do
				End If
			Loop		
			
			str_Msg = "Total de Registros:&nbsp;" & intTotalRegistros
				
		Else
			str_Msg = "Não existem registros para esta consulta"
		End If 
		 
		rds_Usu_Curso.close
		set rds_Usu_Curso = nothing
		%>		
		<tr><td colspan="13" width="100%" height="10"></td></tr>
		
		<tr>
			<td colspan="13" align="center" width="100%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000080" size="2"><b><%=str_Msg%></b></font></td>
		</tr>
		
		<tr><td colspan="13" width="100%" height="10"></td></tr>
		
		</table>			
	</body>	
	<script language="JavaScript" type="text/JavaScript">
		document.imagem1.src = "../../imagens/carregando_limpa.gif"
	</script>
</html>

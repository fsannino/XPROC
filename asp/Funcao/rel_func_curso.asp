<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("opt") = 1 then
   Response.Buffer = TRUE
   Response.ContentType = "application/vnd.ms-excel"
end if

str_Uso = request("chkEmUso")
if str_Uso = "" then
   str_Uso = 0
end if 

str_Desuso = request("chkEmDesuso")  
if str_Desuso = "" then
   str_Desuso = 0
end if   

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso = " AND (FUNC_NEG.FUNE_TX_INDICA_EM_USO = '1' OR FUNC_NEG.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso = " AND FUNC_NEG.FUNE_TX_INDICA_EM_USO = '1' "
   else
      if str_Desuso = 1 then
         str_usoDesuso = " AND FUNC_NEG.FUNE_TX_INDICA_EM_USO = '0' "
	  else
     	 str_usoDesuso = " AND FUNC_NEG.FUNE_TX_INDICA_EM_USO = '2' "
	  end if	 
	end if        	  
end if

str_Opc 				= Request("txtOpc")
str_MegaProcesso 		= request("selMegaProcesso")
str_Modulo				= request("selSubModulo")
str_CdAreaAbrangencia 	= request("selAreaAbrangencia")
str_CdFuncao 			= request("selFuncao")
selG 					= request("selG")
str_Critica 			= request("chkCritica")

compl1 = ""
if str_modulo<>"0" then
	compl1 = " AND FUNC_NEG_SUB.SUMO_NR_CD_SEQUENCIA = " & str_modulo 
end if

if str_CdFuncao <> "0" then
	compl1 = compl1 & " AND FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO = '" & str_CdFuncao  & "'"
end if

str_Sub_Titulo = ""

if selG = "1" then
	compl1 = compl1 & " AND FUNC_NEG.FUNE_TX_TP_FUN_NEG ='G'"
	str_Sub_Titulo =  str_Sub_Titulo & " - Genérica"
else
   selG = 0	
end if

if str_Critica=1 then
	compl1=compl1 + " AND FUNC_NEG.FUNE_TX_INDICA_CRITICA ='1'"
	str_Sub_Titulo = str_Sub_Titulo & " - Crítica"	
else
   str_Critica = 0	
end if

set rs_mega = db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso)

if str_MegaProcesso = 0 then
	IF selG=1 then
		ssql = "SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_TX_TP_FUN_NEG = 'G' " & str_usoDesuso  
	else
		ssql = "SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO > 0 " & str_usoDesuso 
	end if
	ssql = ssql & " ORDER BY FUNE_CD_FUNCAO_NEGOCIO "
else
	if str_modulo <> "0" then
		ssql = ""				
		ssql = ssql & "SELECT DISTINCT FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO, FUNC_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNC_NEG.FUNE_TX_DESC_FUNCAO_NEGOCIO, FUNC_NEG.MEPR_CD_MEGA_PROCESSO, FUNC_NEG.FUNE_TX_TIPO_CLASS "
		ssql = ssql & "FROM FUNCAO_NEGOCIO_SUB_MODULO FUNC_NEG_SUB, FUNCAO_NEGOCIO FUNC_NEG, CURSO_FUNCAO CUR_FUNC "		
		ssql = ssql & "WHERE FUNC_NEG_SUB.FUNE_CD_FUNCAO_NEGOCIO= FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO "
		ssql = ssql & "AND CUR_FUNC.FUNE_CD_FUNCAO_NEGOCIO = FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO "
		ssql = ssql & "AND MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
		ssql = ssql & " ORDER BY FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO "
	else
		ssql = ""		
		ssql = ssql & "SELECT DISTINCT FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO, FUNC_NEG.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNC_NEG.FUNE_TX_DESC_FUNCAO_NEGOCIO, FUNC_NEG.MEPR_CD_MEGA_PROCESSO, FUNC_NEG.FUNE_TX_TIPO_CLASS "
		ssql = ssql & "FROM FUNCAO_NEGOCIO FUNC_NEG, CURSO_FUNCAO CUR_FUNC "	
		ssql = ssql & "WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
		ssql = ssql & " AND CUR_FUNC.FUNE_CD_FUNCAO_NEGOCIO = FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO "
		ssql = ssql & "ORDER BY FUNC_NEG.FUNE_CD_FUNCAO_NEGOCIO "
	end if	
end if
'Response.write ssql & "<br><br><br>" 	
set rsFuncaoSel = db.execute(ssql)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
	</head>	
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" action="" name="frm1">
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
						<td width="27"></td>  <td width="50"></td>
					  <td width="28"></td>
					  <td width="26">&nbsp;</td>
					  <td width="159"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
					
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td>
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center">
					  <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório de Fun&ccedil;&atilde;o x Cursos</font></div>
				  </td>
				</tr>
			</table>
			<p style="margin-top:0; margin-bottom:0">
			<table border="0" width="100%">
			  <tr>
				<td width="27%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega-Processo</font></b></td>
				<td width="37%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Função</font></b></td>
				<td width="36%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Curso</font></b></td>
			  </tr>
			  <%          
			  tem = 0          
			  do until rsFuncaoSel.eof=true			
			  			
				strSQL_Funcao = ""		
				strSQL_Funcao = strSQL_Funcao & "SELECT FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNE_CD_FUNCAO_NEGOCIO "
				strSQL_Funcao = strSQL_Funcao & "FROM FUNCAO_NEGOCIO "
				strSQL_Funcao = strSQL_Funcao & "WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rsFuncaoSel("FUNE_CD_FUNCAO_NEGOCIO") & "'"				
				'Response.write strSQL_Funcao & "<br><br>" 			
				set rsFuncao = DB.EXECUTE(strSQL_Funcao)
				
				mega_atual 	= ""
				curso_atual = ""
							
				do until rsFuncao.eof = true				
					atual1 = rsFuncaoSel("MEPR_CD_MEGA_PROCESSO")
					atual2 = rsFuncaoSel("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
					%>
					<tr>
					<%
					SET RS1 = DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rsFuncaoSel("MEPR_CD_MEGA_PROCESSO"))
					if atual1 <> ant1 then
						strNomeMega = RS1("MEPR_TX_DESC_MEGA_PROCESSO")            
					else
						strNomeMega = ""
					end if
					
					if strNomeMega = "" then
						cor = "white"
					else
						cor = "#CCCCCC"	
					end if			
					%>
					<td width="27%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><b><%=strNomeMega%></b></font></td>
					<%
					if atual2 <> ant2 then
						strNomeFuncao = "<b>" & rsFuncaoSel("FUNE_CD_FUNCAO_NEGOCIO") & "</b> - " & rsFuncaoSel("FUNE_TX_TITULO_FUNCAO_NEGOCIO")            
					else
						strNomeFuncao = ""
					end if
					
					if strNomeFuncao = "" then
						cor = "white"
					else
						cor = "#FFFFDF"	
					end if	
					%>
					<td width="37%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><%=strNomeFuncao%></font></td>
					<%						
					strSQL_Curso = ""
					strSQL_Curso = strSQL_Curso & "SELECT CUR_FUNC.CURS_CD_CURSO , CUR.CURS_TX_NOME_CURSO "
					strSQL_Curso = strSQL_Curso & "FROM CURSO_FUNCAO CUR_FUNC, CURSO CUR "
					strSQL_Curso = strSQL_Curso & "WHERE CUR_FUNC.CURS_CD_CURSO = CUR.CURS_CD_CURSO "
					strSQL_Curso = strSQL_Curso & "AND CUR_FUNC.FUNE_CD_FUNCAO_NEGOCIO='" & rsFuncao("FUNE_CD_FUNCAO_NEGOCIO") & "'"			
					strSQL_Curso = strSQL_Curso & " ORDER BY CUR_FUNC.CURS_CD_CURSO"
					'Response.write 	strSQL_Curso & "<br>"			
					SET rsCurso = db.execute(strSQL_Curso)			          
					
					intContCurso = 0
					do until rsCurso.eof = true		
					
						intContCurso = intContCurso + 1
					
						strNomeCurso = "<b>" & rsCurso("CURS_CD_CURSO") & "</b> - " & rsCurso("CURS_TX_NOME_CURSO")         
					
						if intContCurso = 1 then
							%>							
								<td width="36%" bgcolor="#CCFFCC"><font face="Verdana" size="2" color="#330099"><%=strNomeCurso%></font></td>
							</tr>
							<%
						else
							%>
								<td bgcolor="#FFFFFF">
								<td bgcolor="#FFFFFF">
								<td width="36%" bgcolor="#CCFFCC"><font face="Verdana" size="2" color="#330099"><%=strNomeCurso%></font></td>
							</tr>
							<%
						end if
					 	rsCurso.movenext
					loop
					
					tem = tem + 1
					
					ant1 = rsFuncaoSel("MEPR_CD_MEGA_PROCESSO")
					ant2 = rsFuncaoSel("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
					
					rsFuncao.movenext
					
					atual1 = rsFuncaoSel("MEPR_CD_MEGA_PROCESSO")
					atual2 = rsFuncaoSel("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
		
				  loop	
				  rsFuncaoSel.movenext
			  loop			  
			  
			  on error resume next
			  rsCurso.close
			  set rsCurso = nothing
			  
			  rsFuncao.close
			  set rsFuncao = nothing
			  
			  rsFuncaoSel.close
			  set rsFuncaoSel = nothing
			  err.clear
			  %>          
			</table>
			
			<%if tem = 0 then%>
				<font face="Verdana" size="2" color="#800000"><b>Nenhum Registro encontrado</b></font>
			<%end if%>
		</form>	
	</body>
</html>
<%
if request("selCorte") <> "" then
	Session("Corte") 		= request("selCorte")
end if

strUnidade 				= request("selUnidade")
strDiretoria 			= request("selDiretoria")
strCT					= request("selCT")
strNumRel				= request("pNumRel")
strTituloRel			= request("pTituloRel")

'Response.write "Corte - " & Session("Corte") & "<br>"
'Response.write "strUnidade - " & strUnidade & "<br>"
'Response.write "strDiretoria - " & strDiretoria & "<br>"
'Response.write "strCT - " & strCT & "<br>"
'Response.write "strNumRel - " & strNumRel & "<br>"
'Response.write "strTituloRel - " & strTituloRel & "<br><br>"
'Response.end

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

'******** UNIDADE ******************
strSQLUnidade =  ""
strSQLUnidade = strSQLUnidade & "SELECT UNID.UNID_CD_UNIDADE, UNID.UNID_TX_DESC_UNIDADE, DIR.DIRE_TX_DESC_DIRETORIA, "
strSQLUnidade = strSQLUnidade & "CT.CTRO_TX_NOME_CENTRO_TREINAMENTO, CT.CTRO_CD_CENTRO_TREINAMENTO, UNID.ORLO_CD_ORG_LOT_DIR "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE UNID, GRADE_DIRETORIA DIR, GRADE_CENTRO_TREINAMENTO CT "
strSQLUnidade = strSQLUnidade & "WHERE UNID.ORLO_CD_ORG_LOT_DIR = DIR.ORLO_CD_ORG_LOT "
strSQLUnidade = strSQLUnidade & "AND CT.CTRO_CD_CENTRO_TREINAMENTO = UNID.CTRO_CD_CENTRO_TREINAMENTO "
strSQLUnidade = strSQLUnidade & " AND UNID.CORT_CD_CORTE = " & Session("Corte") 
strSQLUnidade = strSQLUnidade & " AND CT.CORT_CD_CORTE = " & Session("Corte") 

if strDiretoria <> "0" then
	strSQLUnidade = strSQLUnidade & "AND UNID.ORLO_CD_ORG_LOT_DIR = " & strDiretoria 
end if

if strUnidade <> "0" then
	strSQLUnidade = strSQLUnidade & " AND UNID.UNID_CD_UNIDADE = " & strUnidade
end if

if strCT <> "0" then
	strSQLUnidade = strSQLUnidade & " AND UNID.CTRO_CD_CENTRO_TREINAMENTO = " & strCT
end if

strSQLUnidade = strSQLUnidade & " ORDER BY DIR.DIRE_TX_DESC_DIRETORIA, UNID.UNID_TX_DESC_UNIDADE "
'Response.write strSQLUnidade & "<br><br>"
'Response.END

set rdsUnidade= db_banco.execute(strSQLUnidade)
%>
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
							 <a href="gera_consulta_unidade_x_orgao.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;selCT=<%=strCT%>&amp;selCorte=<%=Session("Corte")%>&amp;pTituloRel=<%=strTituloRel%>&amp;pNumRel=<%=strNumRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
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
			  <td height="27">&nbsp;</td>
			</tr>
			<%
			strSQLCorte = ""
			strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
			strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
			strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
			'Response.write strSQLCorte
			'Response.end
			set rsCorte = db_banco.Execute(strSQLCorte)
		 
			if not rsCorte.eof then
				strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & 	rsCorte("CORT_DT_DATA_CORTE")					
			else
				strNomeCorte = ""
			end if
			
			rsCorte.close
			set rsCorte = nothing			 
			%>				
    	</table>		
				
		<table width="803" border="0" cellpadding="2" cellspacing="2">
		  <tr width="500">
		    <td width="10%"></td>
		    <td align="left" colspan="3"><font face="Verdana" size="2" color="#330099"><b>Corte:</b>&nbsp;<%=strNomeCorte%></font></td>		   
	      </tr>
			  <tr>
			    <td></td>
			    <td>&nbsp;</td>
			    <td>&nbsp;</td>
			    <td>&nbsp;</td>
	      </tr>
			  <tr>
			  	<td></td>
				<td width="36%" bgcolor="#D4D0C8"><b><font face="Verdana" size="1" color="#000080">Diretoria</font></b></td>
				<td width="32%" bgcolor="#D4D0C8"><b><font face="Verdana" size="1" color="#000080">Unidade</font></b></td>
				<td width="22%" bgcolor="#D4D0C8"><b><font face="Verdana" size="1" color="#000080">Orgão</font></b></td>
			  </tr>
			  <%          
			  tem = 0          
			  do until rdsUnidade.eof=true			
			  			
				strSQL_Unidade = ""		
				strSQL_Unidade = strSQL_Unidade & "SELECT UNID_CD_UNIDADE, UNID_TX_DESC_UNIDADE "
				strSQL_Unidade = strSQL_Unidade & "FROM GRADE_UNIDADE "
				strSQL_Unidade = strSQL_Unidade & "WHERE UNID_CD_UNIDADE=" & rdsUnidade("UNID_CD_UNIDADE")		
				strSQL_Unidade = strSQL_Unidade & " AND CORT_CD_CORTE = " & Session("Corte") 
				'Response.write strSQL_Unidade & "<br><br>" 			
				set rsUnidade_2 = db_banco.EXECUTE(strSQL_Unidade)
				
				diretoria_atual 	= ""
				orgao_atual = ""
							
				do until rsUnidade_2.eof = true				
					atual1 = rdsUnidade("ORLO_CD_ORG_LOT_DIR")
					atual2 = rdsUnidade("UNID_TX_DESC_UNIDADE")
					%>
					<tr>
					<%
					SET RS1 = db_banco.EXECUTE("SELECT DIRE_TX_DESC_ABREV FROM GRADE_DIRETORIA WHERE ORLO_CD_ORG_LOT = " & rdsUnidade("ORLO_CD_ORG_LOT_DIR") & " ORDER BY DIRE_TX_DESC_ABREV")
					if atual1 <> ant1 then
						strNomeDiretoria = RS1("DIRE_TX_DESC_ABREV")            
					else
						strNomeDiretoria = ""
					end if
					
					if strNomeDiretoria = "" then
						cor = "white"
					else
						cor = "silver"	
					end if			
					%>
					<td></td>
					<td width="36%" bgcolor="<%=cor%>"><font face="Verdana" size="1" color="#000080"><b><%=strNomeDiretoria%></b></font></td>
					<%
					if atual2 <> ant2 then
						strNomeUnidade = rdsUnidade("UNID_TX_DESC_UNIDADE")            
					else
						strNomeUnidade = ""
					end if
					
					if strNomeUnidade = "" then
						cor = "white"
					else
						cor = "#F1F1F6"	'#FFFFDF
					end if	
					%>
					<td width="32%" bgcolor="<%=cor%>"><font face="Verdana" size="1" color="#000080"><b><%=strNomeUnidade%></b></font></td>
					<%				
					strSQLUnidOrgao = ""
					strSQLUnidOrgao = strSQLUnidOrgao & "SELECT ORGAO_MENOR.ORME_SG_ORG_MENOR "
					strSQLUnidOrgao = strSQLUnidOrgao & "FROM GRADE_UNIDADE_ORGAO_MENOR UNID_ORGAO, GRADE_ORGAO_MENOR ORGAO_MENOR "
					strSQLUnidOrgao = strSQLUnidOrgao & "WHERE UNID_ORGAO.ORME_CD_ORG_MENOR = ORGAO_MENOR.ORME_CD_ORG_MENOR "
					'strSQLUnidOrgao = strSQLUnidOrgao & "AND UNID_ORGAO.CORT_CD_CORTE = " & Session("Corte")
					strSQLUnidOrgao = strSQLUnidOrgao & " AND ORGAO_MENOR.CORT_CD_CORTE = " & Session("Corte")
					strSQLUnidOrgao = strSQLUnidOrgao & " AND UNID_ORGAO.UNID_CD_UNIDADE = " & rdsUnidade("UNID_CD_UNIDADE")
					strSQLUnidOrgao = strSQLUnidOrgao & " ORDER BY ORGAO_MENOR.ORME_SG_ORG_MENOR"
					'Response.write strSQLUnidOrgao & "<br>"	
					'Response.end		
					
					SET rsOrgao = db_banco.execute(strSQLUnidOrgao)			          
					
					intContOrgao = 0
					do until rsOrgao.eof = true		
					
						intContOrgao = intContOrgao + 1
					
						strNomeOrgao = rsOrgao("ORME_SG_ORG_MENOR")         
					
						if intContOrgao = 1 then
							%>							
								<td width="22%" bgcolor="#F0F7F0"><font face="Verdana" size="1" color="#000080"><b><%=strNomeOrgao%></b></font></td>
		  					</tr>
							<%
						else
							%>
								<td bgcolor="#FFFFFF"></td>
								<td bgcolor="#FFFFFF"></td>
								<td bgcolor="#FFFFFF"></td>
								<td width="32%" bgcolor="#F0F7F0"><font face="Verdana" size="1" color="#000080"><b><%=strNomeOrgao%></b></font></td>
							</tr>
							<%
						end if
					 	rsOrgao.movenext
					loop
					
					tem = tem + 1
					
					ant1 = rdsUnidade("ORLO_CD_ORG_LOT_DIR")
					ant2 = rdsUnidade("UNID_TX_DESC_UNIDADE")
					
					rsUnidade_2.movenext
					
					atual1 = rdsUnidade("ORLO_CD_ORG_LOT_DIR")
					atual2 = rdsUnidade("UNID_TX_DESC_UNIDADE")
		
				  loop	
				  rdsUnidade.movenext
			  loop			  
			  
			  on error resume next
			  rsOrgao.close
			  set rsOrgao = nothing
			  
			  rsUnidade_2.close
			  set rsUnidade_2 = nothing
			  
			  rdsUnidade.close
			  set rdsUnidade = nothing
			  err.clear
			  
			  if tem = 0 then
			  %>
			  	<tr>
					<td bgcolor="#FFFFFF"></td>
					<td bgcolor="#FFFFFF" colspan="3"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não existe registros para esta consulta.</font></td>
				</tr>			
			  <%
			  end if
			  %>			         
			</table>	
			<p>&nbsp;</p>
		</form>	
	</body>
</html>

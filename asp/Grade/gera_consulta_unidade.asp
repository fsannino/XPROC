<%
Session("Corte") 		= request("selCorte")
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

strSQLUnidade = ""
strSQLUnidade = strSQLUnidade & " SELECT"     
strSQLUnidade = strSQLUnidade & " UNID.UNID_CD_UNIDADE"
strSQLUnidade = strSQLUnidade & " , UNID.UNID_TX_DESC_UNIDADE"
strSQLUnidade = strSQLUnidade & " , DIR.DIRE_TX_DESC_DIRETORIA"
strSQLUnidade = strSQLUnidade & " , CT.CTRO_TX_NOME_CENTRO_TREINAMENTO "
strSQLUnidade = strSQLUnidade & " , CT.CTRO_CD_CENTRO_TREINAMENTO"
strSQLUnidade = strSQLUnidade & " , UNID.ORLO_CD_ORG_LOT_DIR"
strSQLUnidade = strSQLUnidade & " FROM dbo.GRADE_UNIDADE UNID LEFT OUTER JOIN"
strSQLUnidade = strSQLUnidade & " dbo.GRADE_DIRETORIA DIR ON UNID.ORLO_CD_ORG_LOT_DIR = DIR.ORLO_CD_ORG_LOT LEFT OUTER JOIN"
strSQLUnidade = strSQLUnidade & " dbo.GRADE_CENTRO_TREINAMENTO CT ON UNID.CTRO_CD_CENTRO_TREINAMENTO = CT.CTRO_CD_CENTRO_TREINAMENTO"
strSQLUnidade = strSQLUnidade & " WHERE "
strSQLUnidade = strSQLUnidade & " UNID.CORT_CD_CORTE = " & Session("Corte") 
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
			
			/*** Referente aos link ***/
			#Link:visited
			{
				COLOR: #000080;
				font-family: Verdana; 
				font-weight:normal; 
				font-size: 10px;
				TEXT-DECORATION: underline
			}
			#Link:hover
			{
				COLOR: #000080;
				font-family: Verdana; 
				font-weight:normal; 
				font-size: 10px;
				TEXT-DECORATION: none
			}
			#Link
			{
				font-family: Verdana; 
				font-weight:normal; 
				font-size: 10px;
				COLOR: #000080;
				TEXT-DECORATION: underline
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
							 <a href="gera_consulta_unidade.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;selCT=<%=strCT%>&amp;selCorte=<%=Session("Corte")%>&amp;pTituloRel=<%=strTituloRel%>&amp;pNumRel=<%=strNumRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
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
				
		<table width="729" border="0" cellpadding="2" cellspacing="2">
		  <tr width="500">
		    <td></td>
		    <td align="left" colspan="3"><font face="Verdana" size="2" color="#330099"><b>Corte:</b>&nbsp;<%=strNomeCorte%></font></td>		   
	      </tr>
		  <tr width="500">
		    <td></td>
		    <td align="center">&nbsp;</td>
		    <td align="center">&nbsp;</td>
		    <td align="center">&nbsp;</td>
	      </tr>
		  <tr bgcolor="#D4D0C8" width="500">
			<td width="210" bgcolor="#FFFFFF"></td>			    
			<td width="163" align="center">
					<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Diretoria</b></font>
			</td>			
			<td width="147" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Unidade</b></font>
			</td>									
		    <td width="168" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Centro de Treinamento</b></font>
			</td>	     
		  </tr>		
			<%
			'*** INICIALIZAÇÕES ***	
			strCor = "#FFFFFF"
			intTotal = 0			
									
			if not rdsUnidade.eof then									
				do until rdsUnidade.eof = true
														
					if strCor = "#FFFFFF" then
						strCor = "#EAEAEA"
					else
						strCor = "#FFFFFF"
					end if
					%>		
					<tr bgcolor="<%=strCor%>">
						<td width="210" bgcolor="#FFFFFF"></td>		
						<td width="163" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rdsUnidade("DIRE_TX_DESC_DIRETORIA")%></font>
						</td>						
						<td width="147" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">
								<%if request("excel") <> 1 then%>
									<a href="#" id="Link" Onclick="window.open('mostra_unid_orgao_menor.asp?pCdCorte=<%=Session("Corte")%>&pCdDiretoria=<%=rdsUnidade("ORLO_CD_ORG_LOT_DIR")%>&pCdCT=<%=rdsUnidade("CTRO_CD_CENTRO_TREINAMENTO")%>&pCdUnidade=<%=rdsUnidade("UNID_CD_UNIDADE")%>','janela_orgao','toolbar=no,location=no,directories=no,status=no,scrollbars=yes,menubar=no,resizable=no,width=600px,height=500,top=100,left=150');">
										<%=rdsUnidade("UNID_TX_DESC_UNIDADE")%>
									</a>
								<%else%>
									<%=rdsUnidade("UNID_TX_DESC_UNIDADE")%>
								<%end if%>
							</font>
						</td>												
						<td width="168" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rdsUnidade("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></font>
						</td>						
					</tr>					
					<%
					intTotal = intTotal + 1					
					rdsUnidade.movenext					
				loop
				%>
				<tr>
					<td width="210" bgcolor="#FFFFFF"></td>		
					<td height="20"></td>
				</tr>
				<tr>
					<td width="210" bgcolor="#FFFFFF"></td>		
					<td colspan="3">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total de Unidades:</b></font>&nbsp;&nbsp;  
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><%=intTotal%></b></font>				  
					</td>
				</tr>				
				<tr>
					<td width="210" bgcolor="#FFFFFF"></td>		
				  <td height="10"></td>
				</tr>
				<%
			else
			%>
				<tr>
					<td width="210" bgcolor="#FFFFFF"></td>	
					<td colspan="3">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não existe registros para esta consulta.</font>
					</td>
				</tr>					
			<%
			end if
			
			rdsUnidade.close
			set rdsUnidade= nothing			
			%>
	</table>		
	<br>
	<br>
	</body>	
	<%
	db_banco.close
	set db_banco = nothing
	%>		
</html>

<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

intCdCorte = Request("pCdCorte")
intCdDiretoria = Request("pCdDiretoria")
intCdCT = Request("pCdCT")
intCdUnidade = Request("pCdUnidade")

'Response.write "intCdCorte - " & intCdCorte & "<br>"
'Response.write "intCdDiretoria - " & intCdDiretoria & "<br>"
'Response.write "intCdCT - " & intCdCT & "<br>"
'Response.write "intCdUnidade - " & intCdUnidade & "<br><br>"

'******** UNIDADE ******************
strSQLUnidade =  ""
strSQLUnidade = strSQLUnidade & "SELECT UNID.UNID_CD_UNIDADE, UNID.UNID_TX_DESC_UNIDADE, DIR.DIRE_TX_DESC_DIRETORIA, "
strSQLUnidade = strSQLUnidade & "CT.CTRO_TX_NOME_CENTRO_TREINAMENTO, CT.CTRO_CD_CENTRO_TREINAMENTO, UNID.ORLO_CD_ORG_LOT_DIR "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE UNID, GRADE_DIRETORIA DIR, GRADE_CENTRO_TREINAMENTO CT "
strSQLUnidade = strSQLUnidade & "WHERE UNID.ORLO_CD_ORG_LOT_DIR = DIR.ORLO_CD_ORG_LOT "
strSQLUnidade = strSQLUnidade & "AND CT.CTRO_CD_CENTRO_TREINAMENTO = UNID.CTRO_CD_CENTRO_TREINAMENTO "
strSQLUnidade = strSQLUnidade & " AND UNID.CORT_CD_CORTE = " & intCdCorte
strSQLUnidade = strSQLUnidade & " AND UNID.ORLO_CD_ORG_LOT_DIR = " & intCdDiretoria
strSQLUnidade = strSQLUnidade & " AND UNID.CTRO_CD_CENTRO_TREINAMENTO = " & intCdCT
strSQLUnidade = strSQLUnidade & " AND UNID.UNID_CD_UNIDADE = " & intCdUnidade
'Response.write strSQLUnidade & "<br><br>"
'Response.END

set rdsUnidade = db_banco.execute(strSQLUnidade)

if not rdsUnidade.eof then
	strNomeDiretoria = rdsUnidade("DIRE_TX_DESC_DIRETORIA")
	strNomeCT = rdsUnidade("CTRO_TX_NOME_CENTRO_TREINAMENTO")
	strNomeUnidade = rdsUnidade("UNID_TX_DESC_UNIDADE")
			
	'******* PEGA O NOME DO CORTE *****
	strSQLCorte = ""
	strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
	strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
	strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & intCdCorte
	'Response.write strSQLCorte
	'Response.end
	set rsCorte = db_banco.Execute(strSQLCorte)
	
	if not rsCorte.eof then
		strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")					
	else
		strNomeCorte = ""
	end if
	
	rsCorte.close
	set rsCorte = nothing	
	'***** FIM CORTE *****		
else
	strNomeCorte = ""
	strNomeDiretoria = ""
	strNomeCT = ""
	strNomeUnidade = ""
end if

strSQL_OrgaoMenor = ""
strSQL_OrgaoMenor =strSQL_OrgaoMenor & "SELECT ORGAO_MENOR.ORME_SG_ORG_MENOR "
strSQL_OrgaoMenor =strSQL_OrgaoMenor & "FROM GRADE_UNIDADE_ORGAO_MENOR UNID_ORGAO, ORGAO_MENOR "
strSQL_OrgaoMenor =strSQL_OrgaoMenor & "WHERE UNID_ORGAO.ORME_CD_ORG_MENOR = ORGAO_MENOR.ORME_CD_ORG_MENOR "
strSQL_OrgaoMenor =strSQL_OrgaoMenor & "AND UNID_ORGAO.UNID_CD_UNIDADE = " & intCdUnidade
strSQL_OrgaoMenor =strSQL_OrgaoMenor & " ORDER BY ORGAO_MENOR.ORME_SG_ORG_MENOR"

'Response.write strSQL_OrgaoMenor
'Response.end

set rdsOrgaoMenor = db_banco.execute(strSQL_OrgaoMenor) 	
%>
<html>
	<head>
	</head>
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
	
		<table cellspacing="0" cellpadding="0" border="0" width="91%">
			<tr>
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3"><b>Orgãos Associados - Grade de Treinamento</b></font></div>
			  </td>
			</tr>	
	</table>		
	
		<table width="565" border="0" cellpadding="2" cellspacing="2">
		  <tr width="500">
		    <td></td>
		    <td align="left">&nbsp;</td>
		    <td>&nbsp;</td>
	      </tr>
		  <tr width="500">
		    <td width="31"></td>
		    <td width="187" align="left"><font face="Verdana" size="2" color="#330099"><b>Corte:</b></font></td>		   
	      	<td width="327"><font face="Verdana" size="2" color="#330099"><%=strNomeCorte%></font></td>
		  </tr>
		  <tr width="500">
		    <td width="31"></td>
		    <td width="187" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>		   
	      	<td width="327"><font face="Verdana" size="2" color="#330099"><%=strNomeDiretoria%></font></td>
		  </tr>
		   <tr width="500">
		    <td width="31"></td>
		    <td width="187" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:</b></font></td>		   
	      	<td width="327"><font face="Verdana" size="2" color="#330099"><%=strNomeCT%></font></td>
		  </tr>
		    <tr width="500">
		    <td width="31"></td>
		    <td width="187" align="left"><font face="Verdana" size="2" color="#330099"><b>Unidade:</b></font></td>		   
	      	<td width="327"><font face="Verdana" size="2" color="#330099"><%=strNomeUnidade%></font></td>
		  </tr>
   	 </table>
		 
		<table width="563" border="0" cellpadding="2" cellspacing="2">
		  <tr width="500">
		    <td></td>
		    <td>&nbsp;</td>
	       </tr>
		  <tr width="500">
		    <td width="31"></td>
		    <td width="518" align="left" bgcolor="#D4D0C8"><font face="Verdana" size="2" color="#330099"><b>Orgão Menor:</b></font></td>		
		  </tr>
		  			  
			<%
			'*** INICIALIZAÇÕES ***	
			strCor = "#FFFFFF"
			intTotal = 0
			
			if not rdsOrgaoMenor.eof then
				do while not rdsOrgaoMenor.eof
				
					if strCor = "#FFFFFF" then
						strCor = "#EAEAEA"
					else
						strCor = "#FFFFFF"
					end if
					%>
					<tr bgcolor="<%=strCor%>">
						 <td width="31" bgcolor="#FFFFFF"></td>
						<td width="518"><font face="Verdana" size="2" color="#330099"><%=rdsOrgaoMenor("ORME_SG_ORG_MENOR")%></font></td>
					</tr>
					<%
					intTotal = intTotal + 1		
					rdsOrgaoMenor.movenext
				loop
			else
				%>
				<tr>			
					 <td width="31" bgcolor="#FFFFFF"></td>		
				    <td>
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não existe registros para esta consulta.</font>
				    </td>
				</tr>				
				<%
			end if
			%>		  
		 </table>
	</body>
</html>
<br>
<br>
<%
db_banco.close
set db_banco = nothing
%>
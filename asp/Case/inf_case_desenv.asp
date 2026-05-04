<%
strCdDesenv = Request("CdDesenv")

'Response.write "DTPrevRealiz = " & Request("DTPrevRealiz") & "<br>"
'Response.write "DTConclusao = " & Request("DTConclusao") & "<br>"
'Response.write "DTIniTrans = " & Request("DTIniTrans") & "<br>"

if trim(Request("DTPrevRealiz")) <> "" then	
	strDTPrevRealiz = cdate(trim(Request("DTPrevRealiz")))
	strDTPrevRealizMostra = trim(MontaDataHora(trim(Request("DTPrevRealiz")),2))
else
	strDTPrevRealiz = ""
	strDTPrevRealizMostra = "N/A"
end if

if trim(Request("DTConclusao")) <> "" then
	strDTConclusao = cdate(trim(Request("DTConclusao")))
	strDTConclusaoMostra = trim(MontaDataHora(trim(Request("DTConclusao")),2))
else
	strDTConclusao = ""
	strDTConclusaoMostra = "N/A"
end if

if trim(Request("DTIniTrans")) <> "" then	
	strDTIniTrans = cdate(trim(Request("DTIniTrans")))
	strDTIniTransMostra = trim(MontaDataHora(trim(Request("DTIniTrans")),2))
else
	strDTIniTrans = ""
	strDTIniTransMostra = "N/A"
end if

strMsg = trim(Request("msg"))

strDataAtual = date()

strTextoMsg = ""					
if strMsg = "" then		
	if strDTPrevRealiz = "" then
		strTextoMsg = strTextoMsg & " - Não existe Data Prevista de Realização para o referido Desenvolvimento.<br><br>"	
	else
		if strDTPrevRealiz > strDTIniTrans then
			strTextoMsg = strTextoMsg & " - A Data Prevista de Realização do Desenvolvimento é maior do que a Data Início da Transação Case.<br><br>"
		else
			if strDTIniTrans > strDataAtual then
				strTextoMsg = strTextoMsg & " - A Data Início da Transação Case é maior do que a Data de hoje - " & MontaDataHora(trim(strDataAtual),2) & ".<br><br>"
			end if		
		end if
	end if	
	
	if strDTConclusao = "" and strDTPrevRealiz < strDataAtual then
		strTextoMsg = strTextoMsg & " - Não existe Data de Conclusão cadastrada para o referido Desenvolvimento, e a Data Prevista de Realização do Desenvolvimento é menor do que a Data de hoje - " & MontaDataHora(trim(strDataAtual),2) & ".<br><br>"	
	end if		
else	
	strTextoMsg = strMsg
end if

'*****************************************************************************************
'*****************************************************************************************
'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
'*** intDataTime = 1 (DATA E HORA)
'*** intDataTime = 2 (DATA)
'*** intDataTime = 3 (HORA)
'*****************************************************************************************
'*****************************************************************************************
public function MontaDataHora(strData,intDataTime)	

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
	end if
end function
%>
<HTML>
	<HEAD>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<STYLE type=text/css>
			BODY {
				SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
		
			.boton_box
				{
					BORDER-RIGHT: black 1px solid;
					BORDER-TOP: black 1px solid;
					BORDER-COLOR: #330099;
					FONT-WEIGHT: bold;
					FONT-SIZE: 12px;
					WORD-SPACING: 2px;
					TEXT-TRANSFORM: capitalize;
					BORDER-LEFT: black 1px solid;
					COLOR: #330099;
					BORDER-BOTTOM: black 1px solid;
					FONT-STYLE: normal;
					FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif;
					BACKGROUND-COLOR: #FFFFFF;
				}
		</STYLE>			
	</HEAD>	
	<BODY marginheight="0" marginwidth="0" topmargin="0" bottommargin="0">		
		<table width="100%" border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
			<td>				
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr height="10"><td colspan="2"></td></tr>
				  <tr>
					<td td colspan="2" align="center"><font face="Verdana" size="3" color="#330099"><b>Situação do Desenvolvimento</b></font></td>						
				  </tr>
				  <tr height="15"><td colspan="2"></td></tr>
				  <tr>
					<td width="46%"><font face="Verdana" size="2" color="#330099"><b>Código Desenvolvimento:</b></font></td>
					<td width="54%"><font face="Verdana" size="2" color="#330099"><%=strCdDesenv%></font></td>
				  </tr>				 				
				  
				  <tr>	
					<td><font face="Verdana" size="2" color="#330099"><b>Data Prevista Realização:</b></font></td>
					<td><font face="Verdana" size="2" color="#330099"><%=strDTPrevRealizMostra%></font></td>
				  </tr>		 
				 
				  <tr>	
					<td><font face="Verdana" size="2" color="#330099"><b>Data de Conclusão:</b></font></td>
					<td><font face="Verdana" size="2" color="#330099"><%=strDTConclusaoMostra%></font></td>
				  </tr>
				  				  
				  <tr>	
					<td><font face="Verdana" size="2" color="#330099"><b>Data Início da Transação Case:</b></font></td>
					<td><font face="Verdana" size="2" color="#330099"><%=strDTIniTransMostra%></font></td>
				  </tr>
				  <tr height="20"><td colspan="2"></td></tr>
				  <tr><td colspan="2"><font face="Verdana" size="2" color="#330099"><b>Mensagens:</b></font></td></tr>
				  <tr height="10"><td colspan="2"></td></tr>
				  <tr>	
					<td colspan="2"><font face="Verdana" size="2" color="#330099"><%=strTextoMsg%></font></td>
				  </tr>
				</table>				
			</td>
		  </tr>		
		  <tr><td height="20"></td></tr>
		  <tr>
			  <td align="center">
			  	<input type="button" value=" Imprimir " name="btmImprimir" class="boton_box" onClick="javascript:window.print();">			
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;			
				<input type="button" value="  Fechar  " name="btmFechar" class="boton_box" onClick="javascript:window.close();">
			  </td>
		  </tr>
		  <tr><td height="20"></td></tr>		  
		</table>		
	</BODY>
</HTML>
		
					
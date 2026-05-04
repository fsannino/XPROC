<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
Session.LCID = 1046

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strNumRel = Request("pNumRel")
strTituloRel = Request("pTituloRel")

if trim(Request("selCorte")) <> "0" and trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))
end if 

if trim(Request("selDiretoria")) <> "" then
	strdiretoria = trim(Request("selDiretoria"))
else
	strdiretoria = "0"
end if

if trim(Request("selCT")) <> "" then
	strCT = trim(Request("selCT"))
else
	strCT = "0"
end if

if trim(request("selUnidade")) <> "" then
	strUnidade = trim(request("selUnidade"))
else
	strUnidade = "0"
end if

'************ CORTE ****************
strSQLCorte = ""
strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
'Response.write strSQLCorte
'Response.end

set rsCorte = db_banco.Execute(strSQLCorte)
				
'************ DIRETORIA ****************
strSQLDiretoria =  ""
strSQLDiretoria = strSQLDiretoria & "SELECT ORLO_CD_ORG_LOT, DIRE_TX_DESC_DIRETORIA "
strSQLDiretoria = strSQLDiretoria & "FROM GRADE_DIRETORIA "
strSQLDiretoria = strSQLDiretoria & "ORDER BY DIRE_TX_DESC_DIRETORIA "
'Response.WRITE  strSQLDiretoria & "<br><br>"
'Response.END

set rdsDiretoria = db_banco.execute(strSQLDiretoria)

'******** CENTRO DE TREINAMENTO ******************
strSQL_CT = ""
strSQL_CT = strSQL_CT & "SELECT CTRO_CD_CENTRO_TREINAMENTO, CTRO_TX_NOME_CENTRO_TREINAMENTO "
strSQL_CT = strSQL_CT & "FROM GRADE_CENTRO_TREINAMENTO "
strSQL_CT = strSQL_CT & "WHERE CORT_CD_CORTE = " & Session("Corte") 
strSQL_CT = strSQL_CT & " ORDER BY CTRO_TX_NOME_CENTRO_TREINAMENTO"
'Response.write strSQL_CT & "<br><br>"
'Response.END

set rdsCT = db_banco.execute(strSQL_CT)

'******** UNIDADE ******************
strSQLUnidade =  ""
strSQLUnidade = strSQLUnidade & "SELECT UNID_CD_UNIDADE, UNID_TX_DESC_UNIDADE "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE "
strSQLUnidade = strSQLUnidade & " WHERE CORT_CD_CORTE = " & Session("Corte") 

if strDiretoria <> "" and strDiretoria <> "0" then
	strSQLUnidade = strSQLUnidade & " AND ORLO_CD_ORG_LOT_DIR = " & strdiretoria 
'else
	'strSQLUnidade = strSQLUnidade & " AND ORLO_CD_ORG_LOT_DIR = 99999999" 
end if

if strCT <> "" and strCT <> "0" then
	strSQLUnidade = strSQLUnidade & " AND CTRO_CD_CENTRO_TREINAMENTO = " & strCT
end if

strSQLUnidade = strSQLUnidade & " ORDER BY UNID_TX_DESC_UNIDADE "
'Response.write strSQLUnidade & "<br><br>"
'Response.END

set rdsUnidade = db_banco.execute(strSQLUnidade)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>			
	</head>

	<script language="javascript">
	
		var intSpan = 0;
	
		function submet_pagina(strValor, strTipo)
		{										
			if (strTipo == 'Corte')
			{				
				window.location.href = "consulta_unidade.asp?selCorte="+document.frm1.selCorte.value+"&selDiretoria=0&selUnidade=0&selCT=0&pNumRel="+document.frm1.pNumRel.value+"&pTituloRel="+document.frm1.pTituloRel.value;
			}			
			
			if (strTipo == 'Diretoria')
			{
				window.location.href = "consulta_unidade.asp?selCorte="+document.frm1.selCorte.value+"&selDiretoria="+document.frm1.selDiretoria.value+"&selCT="+document.frm1.selCT.value+"&selUnidade=0&pNumRel="+document.frm1.pNumRel.value+"&pTituloRel="+document.frm1.pTituloRel.value;
			}	
			
			if (strTipo == 'CT')
			{
				window.location.href = "consulta_unidade.asp?selCorte="+document.frm1.selCorte.value+"&selDiretoria="+document.frm1.selDiretoria.value+"&selCT="+document.frm1.selCT.value+"&selUnidade="+document.frm1.selUnidade.value+"&pNumRel="+document.frm1.pNumRel.value+"&pTituloRel="+document.frm1.pTituloRel.value;
			}							
		}

		function Confirma()
		{					
			if (document.frm1.pNumRel.value == '1')
			{
				document.frm1.action = "gera_consulta_unidade.asp";
			}
									
			if (document.frm1.pNumRel.value == '2')
			{
				document.frm1.action = "gera_consulta_unidade_x_orgao.asp";
			}
			
			document.frm1.submit();			
		}

		function ver_conteudo(fbox)
		{
			valor=fbox.value;
			tamanho=valor.length;
			str1=valor.slice(tamanho-1,tamanho);
			if (str1!=0 && str1!=1 && str1!=2 && str1!=3 && str1!=4 && str1!=5 && str1!=6 && str1!=7 && str1!=8 && str1!=9){
				fbox.value="";
				str2=valor.slice(0,tamanho-1)
				fbox.value=str2;
			}
		}		
		
</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
	<form method="POST" name="frm1">	
	
		<input type="hidden" name="pNumRel" value="<%=strNumRel%>">
		<input type="hidden" name="pTituloRel" value="<%=strTituloRel%>">
			   
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
					<td width="26"><a href="javascript:Confirma();"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
				  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					 <td width="28"></td>  
						<td width="250"></td>
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
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3"><b>Consulta - <%=strTituloRel%></b></font></div>
			  </td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			</tr>
		  </table>
		  <table border="0" width="849" height="141">					
			<tr>
			  <td height="33"></td>
			  <td height="33" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Corte:</b></font></td>
			  <td height="33" valign="middle" align="left">			  	
				<select name="selCorte" size="1" onchange="javascript:submet_pagina(this.value,'Corte');">							
					<option value="0">== Selecione um Corte ==</option>
					<%										
					do until rsCorte.eof=true											
						if cint(Session("Corte")) = cint(rsCorte("CORT_CD_CORTE")) then
							%>
							<option value="<%=rsCorte("CORT_CD_CORTE")%>" selected><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
							<% 
						else							
							%>
							<option value="<%=rsCorte("CORT_CD_CORTE")%>"><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
							<% 
						end if							
						rsCorte.movenext
					loop
					%>
				</select>	
			  </td>
		    </tr>
			<tr>
			  <td height="32"></td>
			  <td height="32" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>
			  <td height="32" valign="middle" align="left">			  
			  <select size="1" name="selDiretoria" onchange="javascript:submet_pagina(this.value,'Diretoria');">
                <option value="0">== TODAS ==</option>
                <%
					do until rdsDiretoria.eof = true
						  if trim(strdiretoria) = trim(rdsDiretoria("ORLO_CD_ORG_LOT")) then%>
							<option value="<%=replace(rdsDiretoria("ORLO_CD_ORG_LOT"),"&","_")%>" selected><%=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
						<%else%>
							<option value="<%=replace(rdsDiretoria("ORLO_CD_ORG_LOT"),"&","_")%>"><%=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
						<%end if						
						rdsDiretoria.movenext
					loop
					
					rdsDiretoria.close
					set rdsDiretoria = nothing
					%>
              	</select>			
			  </td>
		    </tr>
			
			<tr>
			  <td height="37"></td>
			  <td height="37" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:</b></font></td>
			  <td height="37" valign="middle" align="left">			  	
				<select size="1" name="selCT" onchange="javascript:submet_pagina(this.value,'CT');">
					<option value="0">== TODOS ==</option>
						<%
						do until rdsCT.eof = true
							  if cint(strCT) = cint(rdsCT("CTRO_CD_CENTRO_TREINAMENTO")) then%>
								<option value="<%=rdsCT("CTRO_CD_CENTRO_TREINAMENTO")%>" selected><%=rdsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
							<%else%>
								<option value="<%=rdsCT("CTRO_CD_CENTRO_TREINAMENTO")%>"><%=rdsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
							<%end if						
							rdsCT.movenext
						loop
						
						rdsCT.close
						set rdsCT = nothing
						%>
		      </select>			  
			  </td>
		    </tr> 			
			
			<tr>
			  <td height="26"></td>
			  <td height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Unidade:</b></font></td>
			  <td height="26" valign="middle" align="left">			  	
				<select size="1" name="selUnidade">
					<option value="0">== TODAS ==</option>
						<%
						do until rdsUnidade.eof = true
							  if cint(strUnidade) = cint(rdsUnidade("UNID_CD_UNIDADE")) then%>
								<option value="<%=rdsUnidade("UNID_CD_UNIDADE")%>" selected><%=rdsUnidade("UNID_TX_DESC_UNIDADE")%></option>
							<%else%>
								<option value="<%=rdsUnidade("UNID_CD_UNIDADE")%>"><%=rdsUnidade("UNID_TX_DESC_UNIDADE")%></option>
							<%end if						
							rdsUnidade.movenext
						loop
						
						rdsUnidade.close
						set rdsUnidade = nothing
						%>
			    </select>		 
			  </td>
		    </tr> 			
			
			<tr> 
			  <td width="221" height="1"></td>
			  <td width="214" height="1" valign="middle" align="left"></td>
			  <td height="1" valign="middle" align="left" width="400"> </td>
			</tr>   
	  </table>
</form>

	</body>
</html>

<%
if trim(Session("Conn_String_Cogest_Gravacao")) = "" then
	Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest" 
end if

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

strSQL_Case = ""
strSQL_Case = strSQL_Case & "SELECT DISTINCT CASE_TX_CD_CASE " ', CASE_TX_DESC_CASE "
strSQL_Case = strSQL_Case & "FROM CASE_CONDICAO_TRANS "
strSQL_Case = strSQL_Case & "ORDER BY CASE_TX_CD_CASE"

set rstCase = db_Cogest.execute(strSQL_Case)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<STYLE type=text/css>
			BODY {
				SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
		</STYLE>			
	</head>

	<script>
	
		function Confirma()
		{			
			if (document.frmSelCase.selCase.selectedIndex == 0)
			{
				alert('Para consulta, é necessária a seleção de um case!');
				document.frmSelCase.selCase.focus();
				return;
			}
		
			document.frmSelCase.action = "rel_case_condicao_transacao.asp?case="+document.frmSelCase.selCase.value; 
			document.frmSelCase.submit();
		}
		
	</script>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
	<form method="POST" action="" name="frmSelCase">
		<input type="hidden" name="txtImp" size="20">
		<input type="hidden" name="txtQua" size="20">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
		  <tr>
			<td width="20%" height="20">&nbsp;</td>
			<td width="44%" height="60">&nbsp;</td>
			<td width="36%" valign="top"> 
			  <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
				<tr> 
				  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
					<div align="center">
					  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Case/voltar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Case/avancar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Case/favoritos.gif"></a></div>
				  </td>
				</tr>
				<tr> 
				  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					<div align="center"><a href="javascript:print()"><img border="0" src="../Case/imprimir.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Case/atualizar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					<div align="center"><a href="../../indexA.asp"><img src="../Case/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>
		  <tr bgcolor="#00FF99">
			<td colspan="3" height="20">
			  <table width="625" border="0" align="center">
				<tr>
					<td width="26"><a href="javascript:Confirma()"><img border="0" src="../Case/confirma_f02.gif" width="24" height="24"></a></td>
				  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
				  <td width="26">&nbsp;</td>
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
				<div align="center"><font face="Verdana" color="#330099" size="3">CONSULTA CASE - CONDIÇÃO - TRANSAÇÃO</font></div>
			  </td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			</tr>
		  </table>
		  <table border="0" width="849" height="30">
				  <tr>
					
			  <td width="163" height="26"></td>
					
			  <td width="75" height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Case:</b></font></td>
					
			  <td width="597" height="26" valign="middle" align="left" colspan="2"> 
				  <select size="1" name="selCase">
						<!--<option value="0">== Todos ==</option>-->
						<option value="0">== Selecione um Case ==</option>
						<%do until rstCase.eof = true%>
							<option value="<%=rstCase("CASE_TX_CD_CASE")%>"><%=rstCase("CASE_TX_CD_CASE")%></option>
							<%
							rstCase.movenext
						loop
						%>
					</select>
				</td>					
				  </tr>
			</table>
    </form>	
	<%
	rstCase.close
	set rstCase = nothing
	%>
	</body>
</html>

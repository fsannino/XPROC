<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
Session.LCID = 1046

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

strNumRel = Request("pNumRel")
strTituloRel = Request("pTituloRel")


'************ CORTE ****************
strSQLCorte = ""
strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
'Response.write strSQLCorte
'Response.end

set rsCorte = db_banco.Execute(strSQLCorte)
				
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>			
	</head>

	<script language="javascript">
	
		var intSpan = 0;
	
		function Confirma()
		{					
			if(document.frm1.selCorte1.selectedIndex == 0)
				{
				alert("Selecione o primeiro Corte !");
				document.frm1.selCorte1.focus();
				return;
				}
			if(document.frm1.selCorte2.selectedIndex == 0)
				{
				alert("Selecione o segundo Corte !");
				document.frm1.selCorte2.focus();
				return;
				}
				
			if (document.frm1.pNumRel.value == '1')
			{
				document.frm1.action = "dados_unidade_x_unidade.asp?strTituloRel=<%=strTituloRel%>";
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
		  <table border="0" width="849" height="50">					
			<tr>
			  <td height="26"></td>
			  <td valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Corte:</b></font></td>
			  <td valign="middle" align="left">			  	
				<select name="selCorte1" size="1">							
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
			  <td width="221" height="1"></td>
			  <td width="222" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Corte:</b></font></td>
			  <td height="1" valign="middle" align="left" width="392">
			  <select name="selCorte2" size="1">
                <option value="0">== Selecione um Corte ==</option>
                <%	
					rsCorte.movefirst									
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
              </select> </td>
			</tr>   
	  </table>
</form>

	</body>
</html>

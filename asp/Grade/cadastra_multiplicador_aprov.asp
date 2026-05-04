<%@LANGUAGE="VBSCRIPT"%>
<%
server.scripttimeout = 99999999
response.buffer = false

set db_cogest = server.createobject("ADODB.CONNECTION")
db_cogest.Open Session("Conn_String_Cogest_Gravacao")
db_cogest.CursorLocation = 3

strCurso = Request("selCurso")

if Request("selCorte") <> "" then
	Session("Corte") = Request("selCorte")
end if

strSQLMultiplicador = ""
strSQLMultiplicador = strSQLMultiplicador & "SELECT MULT_CURSO.MULT_NR_CD_ID_MULT, MULT_CURSO.CURS_CD_CURSO, "
strSQLMultiplicador = strSQLMultiplicador & "MULT_CURSO.MULT_TX_APROVEITAMENTO, MULT.MULT_NR_CD_CHAVE, "
strSQLMultiplicador = strSQLMultiplicador & "MULT.MULT_TX_NOME_MULTIPLICADOR "                           
strSQLMultiplicador = strSQLMultiplicador & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
strSQLMultiplicador = strSQLMultiplicador & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
strSQLMultiplicador = strSQLMultiplicador & "AND MULT_CURSO.CURS_CD_CURSO = '" & strCurso & "'"
strSQLMultiplicador = strSQLMultiplicador & " AND MULT_CURSO.CORT_CD_CORTE = " & Session("Corte")
strSQLMultiplicador = strSQLMultiplicador & " AND MULT.MULT_NR_TIPO_MULTIPLICADOR <> 3"
strSQLMultiplicador = strSQLMultiplicador & " ORDER BY MULT.MULT_TX_NOME_MULTIPLICADOR"
'Response.write strSQLMultiplicador
'Response.end

set rsMultiplicador = db_cogest.execute(strSQLMultiplicador)

intTotMult = rsMultiplicador.RecordCount
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	</head>

	<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">
	<form name="frm1" method="POST" action="grava_multiplicador.asp?parAcao=APROVA">
	
	  <input type="hidden" name="txtQuery" size="69" value="<%=strSQLMultiplicador%>">
	  <input type="hidden" name="hdStrCurso" value="<%=strCurso%>">
	
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="border-collapse: collapse" bordercolor="#111111">
		<tr> 
		  <td height="20" colspan="2">&nbsp;</td>
		  <td height="60" colspan="3">&nbsp;</td>
		  <td valign="top" colspan="2"> <table width="150" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC" style="border-collapse: collapse" bordercolor="#111111">
			  <tr>
				<td bgcolor="#330099" width="51" valign="middle" align="right"> 
				  <div align="center"> 
					<p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
				  </div>
				</td>
				<td bgcolor="#330099" width="49" valign="middle" align="center"> 
				  <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
				</td>
				<td bgcolor="#330099" width="50" valign="middle" align="center"> 
				  <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
				</td>
			  </tr>
			  <tr>
				<td bgcolor="#330099" height="12" width="51" valign="middle" align="right"> 
				  <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
				</td>
				<td bgcolor="#330099" height="12" width="49" valign="middle" align="center"> 
				  <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
				</td>
				<td bgcolor="#330099" height="12" width="50" valign="middle" align="center"> 
				  <div align="center"><a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
				</td>
			  </tr>
			</table></td>
		</tr>
		<tr bgcolor="#00FF99"> 
		  <td height="20" width="99">&nbsp; </td>
		  <td height="20" width="77"> <p align="right">
		  <%if intTotMult > 0 then%>
		  	<a href="#" onclick="javascript:document.frm1.submit();"><img border="0" src="../../imagens/confirma_f02.gif"></a>
		  <%end if%>
		  </td>
		  <%if intTotMult > 0 then%>
		  	<td height="20" width="203"><font size="2" face="Verdana" color="#330099"><b>&nbsp;Enviar</b></font>
		  <%end if%> 
		  </td>
		  <td height="20" width="49">&nbsp;</td>
		  <td height="20" width="394">&nbsp;</td>
		  <td height="20" width="94">&nbsp; </td>
		  <td height="20" width="77">&nbsp; </td>
		</tr>
	  </table>			
	  <table border="0" width="88%" height="82">
		<tr>
		  <td width="25%" height="50">&nbsp;</td>
		  <td width="75%" height="50"><font face="Verdana" color="#330099" size="3"><b>Aprovação de Multiplicadores para o curso - <%=strCurso%></b></font></td>
		</tr>
		<tr>
		  <td width="25%" height="26"></td>
		  <td width="75%" height="26"></td>
		</tr>
	  </table>
	  
	  <table border="0" cellpadding="0" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="88%" id="AutoNumber1" height="123">
		 <%
		 if intTotMult > 0 then
		 %>
		 <tr>
			 <td width="16%" height="18"></td>
			 <td width="56%" height="18" bgcolor="#330099"><b><font size="1" face="Verdana" color="#EFEFEF">&nbsp;Multiplicador</font></b></td>
			 <td width="16%" height="18" bgcolor="#330099" align="center"><b><font size="1" face="Verdana" color="#EFEFEF">Chave</font></b></td>
			 <td width="12%" height="18" bgcolor="#330099"><p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="1" face="Verdana" color="#EFEFEF">Aprovado</font></b></td>
		 </tr>
		 <%
		 end if
		 i = 0		
		 do until i = intTotMult
			
			if cor="white" then
				cor="#E2E2E2"
			else
				cor="white"
			end if
			        
			 %>
			 <tr>
				<td width="16%" height="19">&nbsp;</td>
				<td width="56%" height="19" bgcolor="<%=cor%>">&nbsp;<font size="1" face="Verdana" color="#330099"><b><%=rsMultiplicador("MULT_TX_NOME_MULTIPLICADOR")%></b></font></td>
				<td width="16%" height="19" bgcolor="<%=cor%>" align="center"><font size="1" face="Verdana"color="#330099"><b><%=rsMultiplicador("MULT_NR_CD_CHAVE")%></b></font></td>
				<%
					if rsMultiplicador("MULT_TX_APROVEITAMENTO") <> "" then
						checado = "checked"
					else
						checado = ""						
					end if
				%>
				<td width="12%" height="19" bgcolor="<%=cor%>" align="center">					
					<input type="checkbox" name="<%=trim(rsMultiplicador("MULT_NR_CD_ID_MULT")) & "_" & trim(rsMultiplicador("CURS_CD_CURSO"))%>" value="1" <%=checado%>>			   </td>
			 </tr>
			 <%
			i = i + 1
			'ch_ant = rsMultiplicador("USAP_CD_USUARIO")
			rsMultiplicador.movenext
		loop
		
		db_cogest.close
		set db_cogest = nothing
		
		if i = 0 then
			%> 
			<tr>
			  <td width="16%" height="22"></td>
			  <td height="22" colspan="3"><p align="left">
				<font face="Verdana" color="#330099" size="2">Não existem multiplicadores cadastrados para o curso selecionado.</font>
			  </td>
			</tr>
			<%
		else
			%>
			<tr>
			  <td height="30" colspan="4"></td>			  
			</tr>
			<tr>
			  <td width="16%" height="22"></td>
			  <td height="22" colspan="3"><p align="left">
				  <%if intTotMult > 0 then%>
				  	<font face="Verdana" color="#330099" size="2">Total de Multiplicadores:&nbsp;<b><%=intTotMult%></b></font>
				  <%end if%>
			  </td>
			</tr>
			<%	
		end if
		%>
	  </table>
	</form>
	</body>
</html>
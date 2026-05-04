<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
strCentroTreinamento = Request("selCentroTreinamento")
strUnidade = request("selUnidade")
strCurso = request("selCurso")

'response.Write("<p>" & strCentroTreinamento)
'response.Write("<p>" & strUnidade)
'response.Write("<p>" & strCurso)

strSala = Request("pSala")
strAcao = Request("pAcao")
strTurma = Request("pTurma")

'Response.write "Acao - " & strAcao & "<br>"
'Response.write "Sala - " & strSala & "<br>"
'Response.write "Turma - " & strTurma & "<br>"

Response.Expires = 0

if Session("Conn_String_Cogest_Gravacao") = "" then
	Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"
end if

Session.LCID = 1046

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
'db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("JOAO_ Petrobras 2004_v2.mdb")
db_banco.CursorLocation = 3

set db_cogest = Server.CreateObject("ADODB.Connection")
db_cogest.Open Session("Conn_String_Cogest_Gravacao")

if trim(Request("selMegaProcesso")) <> "" then
	strMega = Request("selMegaProcesso")
else
	strMega = ""
end if

if trim(Request("selCurso")) <> "" then
	strCurso = Request("selCurso")
else
	strCurso = ""
end if

'************ CENTRO DE TREINAMENTO ****************
str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT "
str_Sql = str_Sql & " [5 Salas de Treinamento].[Centro de Treinamento] as CdCentroTrein"
str_Sql = str_Sql & " FROM [5 Salas de Treinamento]"
str_Sql = str_Sql & " order by [5 Salas de Treinamento].[Centro de Treinamento]"
set rdsCentroTreinamento = db_banco.execute(str_Sql)

'******** UNIDADE ******************
str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT "
str_Sql = str_Sql & " [9 CT x Unidade].Unidade as CdUnidade"
str_Sql = str_Sql & " , [9 CT x Unidade].[Centro de Treinamento]"
str_Sql = str_Sql & " FROM [9 CT x Unidade]"
str_Sql = str_Sql & " WHERE [9 CT x Unidade].[Centro de Treinamento] ='" & strCentroTreinamento & "'"
str_Sql = str_Sql & " ORDER BY [9 CT x Unidade].Unidade"
set rdsUnidade = db_banco.execute(str_Sql)

'***********  MEGA-PROCESSO *******************
strSQLMega = ""
strSQLMega = strSQLMega & " SELECT DISTINCT "
strSQLMega = strSQLMega & " MEPR_CD_MEGA_PROCESSO "
strSQLMega = strSQLMega & " , MEPR_TX_DESC_MEGA_PROCESSO "
strSQLMega = strSQLMega & " FROM MEGA_PROCESSO "
strSQLMega = strSQLMega & " WHERE MEPR_TX_ABREVIA NOT IN ('TI','GR') "
strSQLMega = strSQLMega & " ORDER BY MEPR_TX_DESC_MEGA_PROCESSO "
set rsMega = db_cogest.execute(strSQLMega)

'***********  CURSO *******************
strSQLCurso = ""
strSQLCurso = strSQLCurso & " SELECT CodCurso "
strSQLCurso = strSQLCurso & " FROM [2 Cursos] "

if strMega <> "" or strMega <> "0" then
	strSQLCurso = strSQLCurso & "WHERE [Mega-Processo] = '" & strMega & "'"
else
	strSQLCurso = strSQLCurso & "WHERE CodCurso = 'ZZZZZZZZZZ' "
end if
strSQLCurso = strSQLCurso & " ORDER BY CodCurso"

'Response.WRITE  strSQLCurso
'Response.END 

set rsCurso = db_banco.execute(strSQLCurso)
'Response.write "strCurso = " & strCurso & "<br><br>"
'***********  MULTIPLICADOR X CURSO *******************
strSQLMultiplicador = ""
strSQLMultiplicador = strSQLMultiplicador & "SELECT Multiplicador as cd_Multiplicador "
strSQLMultiplicador = strSQLMultiplicador & "FROM [4 Multiplicadores x curso] "

if strCurso <> "" or strCurso <> "0" then
	strSQLMultiplicador = strSQLMultiplicador & "WHERE CodCurso ='" & strCurso & "'"
else
	strSQLMultiplicador = strSQLMultiplicador & "WHERE CodCurso = 'ZZZZZZZZZZ'"
end if
strSQLMultiplicador = strSQLMultiplicador & " ORDER BY Multiplicador "

'Response.WRITE strSQLMultiplicador
'Response.END 

set rsMultiplicador = db_banco.execute(strSQLMultiplicador)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>			
	</head>

	<script language="javascript">
	
		function submet_pagina(strValor, strTipo)
		{	
			//alert(strValor + ' - ' + strTipo);
			
			if (strTipo == 'CentTrei')
			{
				window.location.href = "consulta_geral.asp?selCentroTreinamento="+document.frm1.selCentroTreinamento.value+"&selUnidade=0";
			}
			
			if (strTipo == 'Mega')
			{
				//alert("consulta_geral.asp?selCentroTreinamento="+document.frm1.selCentroTreinamento.value+"&selUnidade="+document.frm1.selUnidade.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value)
				window.location.href = "consulta_geral.asp?selCentroTreinamento="+document.frm1.selCentroTreinamento.value+"&selUnidade="+document.frm1.selUnidade.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value
			}
			
			//if (strTipo == 'Curso')
			//{
				//window.location.href = "consulta_geral.asp?selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selCurso="+strValor;
			//}				
		}

		function Confirma()
		{		
			var intOndaAnt	= 0;
			var intEaD		= 0;
			var intDescentr = 0;
			var intInLoco	= 0;
			
			if (document.frm1.chkOndaAntecipada.checked == true)
			{
				intOndaAnt = 1;
			}
		
			if (document.frm1.chkEaD.checked == true)
			{
				intEaD = 1;
			}
			
			if (document.frm1.chkDescentralizado.checked == true)
			{
				intDescentr = 1;
			}
			
			if (document.frm1.chkInLoco.checked == true)
			{
				intInLoco = 1;
			}
			
			//alert('intOndaAnt - ' + intOndaAnt);
			//alert('intEaD - ' + intEaD);
			//alert('intDescentr - ' + intDescentr);
			//alert('intInLoco - ' + intInLoco);
						
			document.frm1.action = "gera_consulta_turma.asp?pOndaAnt="+intOndaAnt+"&pEaD="+intEaD+"&pDescentr="+intDescentr+"&pInLoco="+intInLoco;
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
					<td width="26"><a href="javascript:Confirma();"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
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
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3">Consulta de Turmas - Grade de Treinamento</font></div>
			  </td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			</tr>
		  </table>
		  <table border="0" width="849" height="350">
			<tr>
			  <td height="33"></td>
			  <td height="33" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento: </b></font></td>
			  <td height="33" valign="middle" align="left">
			  <select size="1" name="selCentroTreinamento" onchange="javascript:submet_pagina(this.value,'CentTrei');">
                <option value="0">== Selecione o Centro de Treinamento ==</option>
                <%
					do until rdsCentroTreinamento.eof = true
						  if trim(strCentroTreinamento) = trim(rdsCentroTreinamento("CdCentroTrein")) then%>
							<option value="<%=rdsCentroTreinamento("CdCentroTrein")%>" selected><%=rdsCentroTreinamento("CdCentroTrein")%></option>
						<%else%>
							<option value="<%=rdsCentroTreinamento("CdCentroTrein")%>"><%=rdsCentroTreinamento("CdCentroTrein")%></option>
						<%end if						
						rdsCentroTreinamento.movenext
					loop
					
					rdsCentroTreinamento.close
					set rdsCentroTreinamento = nothing
					%>
              </select></td>
		    </tr>
			<tr>
			  <td height="33"></td>
			  <td height="33" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Unidade: </b></font></td>
			  <td height="33" valign="middle" align="left">
			  	<select size="1" name="selUnidade">
					<option value="0">== Selecione a Unidade ==</option>
						<%
						do until rdsUnidade.eof = true
							  if trim(strUnidade) = trim(rdsUnidade("CdUnidade")) then%>
								<option value="<%=rdsUnidade("CdUnidade")%>" selected><%=rdsUnidade("CdUnidade")%></option>
							<%else%>
								<option value="<%=rdsUnidade("CdUnidade")%>"><%=rdsUnidade("CdUnidade")%></option>
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
			  <td width="116" height="33"></td>
			  <td width="209" height="33" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo:</b></font></td>
			  <td width="504" height="33" valign="middle" align="left"> 
				<select size="1" name="selMegaProcesso" onchange="javascript:submet_pagina(this.value,'Mega');">
				  <option value="0">== Selecione o Mega-Processo ==</option>
					<%
					do until rsMega.eof = true
						  if trim(strMega) = trim(rsMega("MEPR_TX_DESC_MEGA_PROCESSO")) then%>
							  <option value="<%=rsMega("MEPR_TX_DESC_MEGA_PROCESSO")%>" selected><%=rsMega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
						  <%else%>
							  <option value="<%=rsMega("MEPR_TX_DESC_MEGA_PROCESSO")%>"><%=rsMega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
						  <%end if						
						rsMega.movenext
					loop
					
					rsMega.close
					set rsMega = nothing
					%>
			    </select>			  
			  </td>
			</tr> 
			
			<tr> 
			  <td width="116" height="31"></td>
			  <td width="209" height="31" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso:</b></font></td>
			  <td height="31" valign="middle" align="left" width="504">
				<select size="1" name="selCurso">
				  <option value="0">== Selecione o Curso ==</option>
					<%					
					do until rsCurso.eof = true
					 	 if trim(strCurso) = trim(rsCurso("CodCurso")) then%>
							  <option value="<%=rsCurso("CodCurso")%>" selected><%=rsCurso("CodCurso")%></option>
						  <%else%>
							  <option value="<%=rsCurso("CodCurso")%>"><%=rsCurso("CodCurso")%></option>
						  <%end if	
						
						rsCurso.movenext
					loop
					
					rsCurso.close
					set rsCurso = nothing
					%>
			    </select>			  
			  </td>
			</tr>			
			
			<tr>
			  <td height="31"></td>
			  <td height="31" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Onda antecipada: </b></font></td>
			  <td height="31" valign="middle" align="left"><input type="checkbox" name="chkOndaAntecipada"></td>
		    </tr>
			<tr> 
			  <td width="116" height="25"></td>
			  <td width="209" height="25" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b> EaD </b></font></td>
			  <td height="25" valign="middle" align="left" width="504"><input type="checkbox" name="chkEaD"></td>
			</tr>			
			
			<tr> 
			  <td width="116" height="29"></td>
			  <td width="209" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Descentralizado</b></font></td>
			  <td height="29" valign="middle" align="left" width="504"><input type="checkbox" name="chkDescentralizado"></td>
			</tr>   
			
			<tr> 
			  <td width="116" height="27"></td>
			  <td width="209" height="27" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>In Loco :</b></font></td>
			  <td height="27" valign="middle" align="left" width="504"><input type="checkbox" name="chkInLoco"></td>
			</tr>   
			
			<tr> 
			  <td width="116" height="1"></td>
			  <td width="209" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data Início:</b></font></td>
			  <td height="1" valign="middle" align="left" width="504"> 
			  	<input type="text" name="txtDtInicio" maxlength="10" size="10">
				<a href="javascript:show_calendar(true,'frmCadTurma.txtDtInicio','DD/MM/YYYY')"><img src="../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>
			  </td>
			</tr>   
			
			<tr> 
			  <td width="116" height="25"></td>
			  <td width="209" height="25" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Carga Horária:</b></font></td>
			  <td height="25" valign="middle" align="left" width="504"> 
			  	<input type="radio" name="rdCargaHor" value="1"><font face="Verdana" size="2" color="#330099">4 Horas&nbsp;</font>
				<input type="radio" name="rdCargaHor" value="2"><font face="Verdana" size="2" color="#330099">8 Horas&nbsp;</font>
				<input type="radio" name="rdCargaHor" value="3"><font face="Verdana" size="2" color="#330099">16 Horas&nbsp;</font>
				<input type="radio" name="rdCargaHor" value="4"><font face="Verdana" size="2" color="#330099">32 Horas&nbsp;</font>
			  	<input type="radio" name="rdCargaHor" value="5"><font face="Verdana" size="2" color="#330099">40 Horas</font>			  
			 </td>
			</tr>   
			
			<tr> 
			  <td width="116" height="1"></td>
			  <td width="209" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data Fim:</b></font></td>
			  <td height="1" valign="middle" align="left" width="504"> 
			  	<input type="text" name="txtDtFim" maxlength="10" size="10">
				<a href="javascript:show_calendar(true,'frmCadTurma.txtDtFim','DD/MM/YYYY')"><img src="../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>
			  </td>
			</tr>   
			
			<tr> 
			  <td width="116" height="1"></td>
			  <td width="209" height="1" valign="middle" align="left"></td>
			  <td height="1" valign="middle" align="left" width="504"> </td>
			</tr>   
	  </table>
</form>
	</body>
</html>

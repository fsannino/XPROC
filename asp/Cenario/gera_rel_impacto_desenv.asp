<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

strSelDesenv = trim(request("selDesenvolvimento")) 
strTxtDesenv = trim(request("txtDesenvolvimento")) 

if strSelDesenv <> "0" then
	strDesenv = strSelDesenv
elseif strTxtDesenv <> "" then
	strDesenv = strTxtDesenv
end if

strAbreviaMegaprocesso = left(strDesenv,2)

'*** DESENVOLVIMENTO ****	
SQL_DesenvSel = ""
SQL_DesenvSel = SQL_DesenvSel & "SELECT DESE_TX_DESC_DESENVOLVIMENTO "
SQL_DesenvSel = SQL_DesenvSel & "FROM DESENVOLVIMENTO "
SQL_DesenvSel = SQL_DesenvSel & "WHERE DESE_CD_DESENVOLVIMENTO ='" & strDesenv & "'"

set rsDesenvSel = db.execute(SQL_DesenvSel)	
	
if not rsDesenvSel.eof then
	strDesenvolvimentoSel =  ucase(strDesenv) & " - " & rsDesenvSel("DESE_TX_DESC_DESENVOLVIMENTO")
else
	strDesenvolvimentoSel = ucase(strDesenv)
	msgErro = ", pois o Desenvolvimento informado não foi localizado pelo sistema!"
end if
rsDesenvSel.close
set rsDesenvSel = nothing	
'***

'*** MEGA PROCESSO ****	
SQL_MegaProc = ""
SQL_MegaProc = SQL_MegaProc & "SELECT MEPR_TX_DESC_MEGA_PROCESSO "
SQL_MegaProc = SQL_MegaProc & "FROM MEGA_PROCESSO "
SQL_MegaProc = SQL_MegaProc & "WHERE MEPR_TX_ABREVIA ='" & strAbreviaMegaprocesso & "'"
set rsMegaProc = db.execute(SQL_MegaProc)	
	
if not rsMegaProc.eof then
	strNomeMegaProc = rsMegaProc("MEPR_TX_DESC_MEGA_PROCESSO")
else
	strNomeMegaProc = strAbreviaMegaprocesso & " - Não cadastrado"
end if
rsMegaProc.close
set rsMegaProc = nothing	
'***

'*** CENÁRIOS PAR O DESENVOLVIMENTO SELECIONADO ***
SQL_Desenv = ""
SQL_Desenv = SQL_Desenv & "SELECT 	CT.cena_cd_cenario, "
SQL_Desenv = SQL_Desenv & "CT.mepr_cd_mega_processo, "
SQL_Desenv = SQL_Desenv & "CT.cetr_nr_sequencia, "
SQL_Desenv = SQL_Desenv & "CT.cena_nr_sequencia_trans, "
SQL_Desenv = SQL_Desenv & "CT.cetr_tx_desc_transacao "
SQL_Desenv = SQL_Desenv & "FROM cenario_transacao CT, desenvolvimento DESENV, cenario CENA "
SQL_Desenv = SQL_Desenv & "WHERE CT.dese_cd_desenvolvimento = DESENV.dese_cd_desenvolvimento "
SQL_Desenv = SQL_Desenv & "AND CT.cena_cd_cenario = CENA.cena_cd_cenario "
SQL_Desenv = SQL_Desenv & "AND CT.dese_cd_desenvolvimento ='" & strDesenv & "'"
set rsDesenv = db.execute(SQL_Desenv)
%>
<html>
	<head>
		<STYLE type=text/css>
			BODY {
				SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
		</STYLE>
		<style>
			a {text-decoration:none;}
			a:hover {text-decoration:underline;}
		</style>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	</head>
	
	<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
	<form name="frm1" method="POST" action="">
	<%if request("excel") <> 1 then%>
		 <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
			<tr>
			  <td width="160" height="20" colspan="2">&nbsp;</td>
			  <td width="346" height="60" colspan="3">&nbsp;</td>
			  <td width="330" valign="top" colspan="2"> 
				<table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
				<tr> 
				  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
					<div align="center">
					  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a></div>
				  </td>
				  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
				  </td>
				  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
				  </td>
				</tr>
				<tr> 
				  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					<div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					  <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>
		  <tr bgcolor="#00FF99"> 
			<td height="20" width="155">&nbsp; 
			  
			</td>
			<td colspan="2" height="20" width="31">&nbsp; 
			  
			</td>
			<td height="20" width="244">&nbsp; 
			  
			</td>
			<td colspan="2" height="20" width="112"><a href="gera_rel_impacto_desenv.asp?excel=1&amp;selDesenvolvimento=0&amp;txtDesenvolvimento=<%=strDesenv%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a> 
			  
			</td>
			<td height="20" width="290">&nbsp; 
			  
			</td>
		  </tr>
		</table>
	<%end if%>
	
	  <p align="center"><font color="#330099" face="Verdana" size="3">Relatório de Cenários Impactados por Desenvolvimentos</font></p>
	
	  <table border="0" width="100%">	  	
		<tr>
		  <td width="9%" bgcolor="#FFFFFF">
		  <td width="91%">
			<p style="margin-top: 0; margin-bottom: 0">&nbsp;<font face="Verdana" size="2" color="#330099"><b>Desenvolvimento:&nbsp;<%=strDesenvolvimentoSel%></b></font></p>
		  </td>
		</tr>
		<tr>
		  <td width="9%" bgcolor="#FFFFFF">
		  <td width="91%">
			<p style="margin-top: 0; margin-bottom: 0">&nbsp;<font face="Verdana" size="2" color="#330099"><b>Mega Processo:&nbsp;<%=strNomeMegaProc%></b></font></p>
		  </td>
		</tr>
		<tr><td colspan="4" height="10"></td>
		</tr>	
	  </table>
	  <%
	  tem = 0
	  if not rsDesenv.eof then%>	  
	  
	  <table border="0" width="800" cellspacing="1" cellpadding="0" align="center">
		<tr> 		 
		  <td colspan="4"></td>
		</tr>		
		<tr> 		 
		  <td width="95" bgcolor="#330099" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Código</font></b></td>
		  <td width="629" bgcolor="#330099" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Título</font></b></td> 	   
		</tr>  
		<%
		end if
		
		intCenarioTotal = 0
		do until rsDesenv.eof = true	
			tem = tem + 1		
			
			intCenarioTotal = intCenarioTotal + 1
						
			SQL_Cenario = ""
			SQL_Cenario = "SELECT * FROM " & Session("PREFIXO") & "CENARIO "
			SQL_Cenario = SQL_Cenario & "WHERE cena_cd_cenario ='" & trim(rsDesenv("cena_cd_cenario")) & "'"
			set rsCenario = db.execute(SQL_Cenario)
			
			do until rsCenario.eof = true
				if cor = "white" then
					cor="#E4E4E4"
				else
					cor="white"
				end if					  
				%>
				<tr> 				 
				  <td width="95" height="20" align="center" bgcolor="<%=cor%>"> 
					<font face="Verdana" size="1"><%=rsCenario("CENA_CD_CENARIO")%></font>	  
				  </td>     
				  <td width="629" align="left" bgcolor="<%=cor%>"> 
					<font face="Verdana" size="1"><%=Ucase(rsCenario("CENA_TX_TITULO_CENARIO"))%></font>
				  </td>    
				</tr>
				<%			
				rsCenario.movenext
			loop	
			rsCenario.close
			set rsCenario = nothing
								
			rsDesenv.movenext			
		loop
					
		rsDesenv.close
		set rsDesenv = nothing
		
		if tem <> 0 then
		%>			 	
	  		</table>			
			<table border="0" width="800">
				<tr><td colspan="3" height="10"></td></tr>
				<tr> 
				  <td width="12%" bgcolor="#FFFFFF"></td> 				    
				  <td width="88%" height="20" colspan="2" align="left"> 
					<font face="Verdana" size="2" color="#330099"><b>Total de Cenários:&nbsp;<%=intCenarioTotal%></b></font>	  
				  </td>				      
				</tr>
			</table>
			<BR>			
		<%
		else
		%>
			<BR>			
		<%
		end if
		%>                
	</form>
	<%if tem = 0 then%>
		<p style="margin-top: 0; margin-bottom: 0" align="center"><font color="#800000" face="Verdana" size="2"><b>Não existe nenhum cenário cadastrado para a seleção<%=msgErro%></b></font></p>
	<%end if%>
	</body>
</html>
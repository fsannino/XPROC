<%@LANGUAGE="VBSCRIPT"%> 
<%
Session.TimeOut = 120
Response.Expires = 0
Session.LCID = 1046

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
		
if trim(Request("selCorte")) <> "" then
	Session("Corte") = cint(Request("selCorte"))	
else
	Session("Corte") = 0	
end if 
	
strSQLCorte = ""
strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
strSQLCorte = strSQLCorte & "ORDER BY CORT_DT_DATA_CORTE DESC "
'Response.write strSQLCorte
'Response.end
set rsCorte = db_banco.Execute(strSQLCorte)
%>
<script language="javascript">
	//*** ESTA VARIÁVEL RECEBERÁ A CATEGORIA DO USUÁRIO PARA MONTAR O MENU
	var str_CategoriaUsuario = "<%=Session("CatUsu")%>";	
</script>
<%
    ls_Script = "<script language=""JavaScript"" src=""Templates/js/grade/indexMenu.js""></script>"	
%>  
<html>
<head>	
	<title>SINERGIA # XPROC # Sistema de Cadastro</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<LINK REL="SHORTCUT ICON" href="http://JOAO/XPROC/imagens/Wrench.ico">
	<script language="JavaScript">
	<!--
	function MM_reloadPage(init) {  //reloads the window if Nav4 resized
	  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
		document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
	  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
	}
	MM_reloadPage(true);
	// -->
	</script>
		<style type="text/css">
			<!--
			
			BODY 
			{
			SCROLLBAR-HIGHLIGHT-COLOR: white; 
			SCROLLBAR-SHADOW-COLOR: white; 
			SCROLLBAR-ARROW-COLOR: yellow; 
			SCROLLBAR-BASE-COLOR: #003399; 
			scrollbar-3d-light-color: White
			}
		
			#Link:visited
			{
				COLOR: #330099;
				font-family: Verdana; 
				font-weight:bold; 
				font-size: 12px;	
				TEXT-DECORATION: none			
			}			
			
			#Link:hover
			{
				COLOR: #D4D0C8;
				font-family: Verdana; 
				font-weight:bold; 
				font-size: 12px;
				TEXT-DECORATION: underline
			}			
			
			#Link
			{
				font-family: Verdana; 
				font-weight:bold; 
				font-size: 12px;
				COLOR: #330099;
				TEXT-DECORATION: none				
			}
			
			.style6 
			{
				font-size: 8pt;
				font-family: Arial, Helvetica, sans-serif;
				font-weight: bold;
			}
			-->
		</style>
</head>
	<script>
		function Fecha()
		{
			alert('Você está saindo do ambiente do Aplicativo...Obrigado por utilizar o X-PROC');
		}
		
		function mover()
		{
			window.moveTo(0,0);
		}
		
		function ver_tecla()
		{
			var a = event.keyCode;
			if(a==16){
			alert('Propriedade SINERGIA @ 2003');
			alert(event.width);
			}
		}
		
		function submet_pagina(strValor)
		{					
			//alert(strValor);
			window.location.href = "indexA_grade.asp?selCorte="+strValor;												
		}
	</script>
	<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:mover()" onKeyDown="javascript:ver_tecla()">
		<%=ls_Script%>
		<script type= "text/javascript" language= "JavaScript">
		<!--
		goMenus();
		//-->
		</script>
		
		<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
		  <tr> 
			<td width="136" height="20"><%'=session("MegaProcesso")%><font color="#FFFFFF">&nbsp;</font><%'=Application("Totalvisitas")%></td>
			<td width="362" height="60" valign="middle" colspan="2"> <p align="left"> 
				<%'="aaa " & Session("AcessoUsuario")%>
				<font color="#FFFFFF"><%'=Application("Datainicial")%> <b><font size="1" face="Arial">
				<% if Session("CategoriaUsu") = "indexQ.htm" then %>
					</font></b><span class="style6"><%=Session("Conn_String_Cogest_Gravacao")%></span><b><font size="1" face="Arial">
				<% end if %>
					</font>
					<%'=Session("CategoriaUsu")%>
					</b> </font></p>
			</td>
			<td width="279" valign="top"> 
			  <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
				<tr> 
				  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
					<div align="center"> 
					  <p align="center"><a href="JavaScript:history.back()"><img src="../xproc/imagens/voltar.gif" width="30" height="30" border="0"></a>
					</div>
				  </td>
				  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.forward()"><img src="../xproc/imagens/avancar.gif" width="30" height="30" border="0"></a></div>
				  </td>
				  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000ws10.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Proc')"><img src="../xproc/imagens/favoritos.gif" width="30" height="30" border="0"></a></div>
				  </td>
				</tr>
				<tr> 
				  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					<div align="center"><a href="javascript:print()"><img src="../xproc/imagens/imprimir.gif" width="30" height="30" border="0"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.go()"><img src="../xproc/imagens/atualizar.gif" width="30" height="30" border="0"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					<p align="center">
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>
		  <tr bgcolor="#F1F1F1"> 
			<td colspan="2" height="20" width="550"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu&aacute;rio 
			  : <%=Session("CdUsuario")%></font></b>
			  <%'if session("MegaProcesso")<>0 then
			'set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & session("MegaProcesso"))
			'VALOR=RS("MEPR_TX_DESC_MEGA_PROCESSO")
			%>
			<!--<font face="Verdana" size="2"><b>Mega-Processo Atual : <%'=valor%></b></font>-->
			<%'end if%></td>
			<td colspan="2" height="20" width="231"><font face="Verdana" size="2"><b></b></font></td>
		  </tr>
		</table>
		<table border="0" width="91%">
		  <tr> 
			<td width="1%"></td>
			<td width="98%" align="left">&nbsp; 
			  <p><!--<img src="../xproc/imagens/fundoXProc.jpg" width="692" height="360" border="0"> -->
			  
			  
			  
			  <!--------------------------------------------------------------------------------------------------->
			  
		<form method="POST" name="frm1">			
					
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			  <td height="10"></td>
			</tr>
			
			<tr>
			  <td align="right" width="100%">	
			  
			  		<%if not rsCorte.eof then%>
			  		 
						<font color="#31009C" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>CORTE:</b></font>					
						<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						  <select name="selCorte" onchange="javascript:submet_pagina(this.value);">							
							<%			
							intCount = 1			
							do until rsCorte.eof=true
							
								if Session("Corte") = 0 and intCount = 1 then
									Session("Corte") = rsCorte("CORT_CD_CORTE")								
								end if
								
								if Session("Corte") = rsCorte("CORT_CD_CORTE") then
									%>
									<option value="<%=rsCorte("CORT_CD_CORTE")%>" selected><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
									<% 
								else							
									%>
									<option value="<%=rsCorte("CORT_CD_CORTE")%>"><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
									<% 
								end if
								intCount = intCount + 1
								rsCorte.movenext
							loop
							%>
						</select>
						</font>				
					<%
					end if
					
					rsCorte.close
					set rsCorte = nothing
				%>
			  </td>
		    </tr>
			<tr>
			  <td height="10"></td>
			</tr>
		  </table>
			  
		  <table border="0" width="883" height="130">						
			<tr>
			  <td width="664" align="left" colspan="2"><img src="asp/Grade/imagens/fundoGrade.gif"></td> 			  
			</tr> 
			<tr>
			  <td width="664"></td>			 
			  <td width="209" height="21" align="left" valign="middle"></td>			 
		    </tr>			   
		  </table>
	</form>
			  
			  
			  
			  <!--------------------------------------------------------------------------------------------------->  
			  
			  
			  
			</td>			
		  </tr>
		  <tr> 
			<td width="1%">
			  <p style="margin-top: 0; margin-bottom: 0"></td>
			<td width="98%">&nbsp; </td>			
		  </tr>
	</table>
	</body>
</html>
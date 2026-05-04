<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))
'Response.write strAcao & "<br>"

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

if strAcao = "A" then	
	
	strCdCorte = trim(Request("selCorte"))
	'Response.write strCdCorte & "<br>"

	strSQLAltCorte = ""
	strSQLAltCorte = strSQLAltCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
	strSQLAltCorte = strSQLAltCorte & "FROM GRADE_CORTE " 
	strSQLAltCorte = strSQLAltCorte & "WHERE CORT_CD_CORTE = "& strCdCorte

	Set rdsAltCorte = db_banco.Execute(strSQLAltCorte)			
	
	if not rdsAltCorte.EOF then	
		strNomeCorte	= rdsAltCorte("CORT_TX_DESC_CORTE")				
		if trim(rdsAltCorte("CORT_DT_DATA_CORTE")) <> "" then
			strDtCorte = MontaDataHora(trim(rdsAltCorte("CORT_DT_DATA_CORTE")),2)
		else
			strDtCorte = ""
		end if
	end if
	
	rdsAltCorte.close
	set rdsAltCorte = nothing	
end if

public function MontaDataHora(strData,intDataTime)

	'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
	'*** intDataTime = 1 (DATA E HORA)
	'*** intDataTime = 2 (DATA)
	'*** intDataTime = 3 (HORA)
	'*** intDataTime = 4 (FORMATO DE BANCO)
	'*** intDataTime = 5 (FORMATO DE BANCO - DIA E MÊS)

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
	elseif cint(intDataTime) = 4 then
		MontaDataHora = strMes & "/" & strDia & "/" & year(strData)
	elseif cint(intDataTime) = 5 then
		MontaDataHora = strDia & "/" & strMes
	end if
end function
%>
<html>
	<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="javascript" src="js/digite-cal.js"></script>
<script language="javascript">
			function Confirma()
			{				
				if(document.frmCadCorte.txtNomeCorte.value == "")
				{
					alert("É necessário o preenchimento do campo DESCRIÇÃO DO CORTE!");
					document.frmCadCorte.txtNomeCorte.focus();
					return;
				}					
							
				if(document.frmCadCorte.txtDtCorte.value == "")
				{
					alert("É necessário o preenchimento do campo DATA DO CORTE!");
					document.frmCadCorte.txtDtCorte.focus();
					return;
				}					
								
				if (document.frmCadCorte.parAcao.value == 'I')
				{ 
					document.frmCadCorte.action="grava_corte.asp";
					MM_setTextOfLayer('aviso','','%3Cb%3E%0D%0A%3Cfont face=%22Verdana%22 color=%22red%22 size=%222%22%3E%0D%0AAGUARDE, CRIANDO DADOS DE CORTE...%0D%0A%3C/font%3E%0D%0A%3Cb%3E')					
				}
				
				if (document.frmCadCorte.parAcao.value == 'A')
				
				{ 				
					document.frmCadCorte.action="grava_corte.asp";
				}
				
				document.frmCadCorte.submit();				
			}
		</script>
<script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_setTextOfLayer(objName,x,newText) { //v3.0
  if ((obj=MM_findObj(objName))!=null) with (obj)
    if (navigator.appName=='Netscape') {document.write(unescape(newText)); document.close();}
    else innerHTML = unescape(newText);
}
//-->
</script>
</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		
<form method="POST" name="frmCadCorte">
  <input type="hidden" value="<%=strAcao%>" name="parAcao"> 			
			
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
						<td width="24"><a href="javascript:Confirma();"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
					  <td width="46"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
					  <td width="21">&nbsp;</td>
					  <td width="177"></td>
						<td width="30"></td>  
						<td width="234"></td>
					    <td width="9"></td>
					  <td width="8">&nbsp;</td>
					  <td width="38"></td>
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Corte - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="695" height="168">
			  
			  	<tr>
			  	  <td height="20"></td>
			  	  <td height="20" valign="middle" align="center" colspan="2">
				  	<%if strGrava = "GravaTurma" then%>
				  	<font face="Verdana" color="#FE5A31" size="2"><b><%=strMSG%></b></font>
					<%end if%>
				  </td>			  	  
		  	    </tr>
			  	<tr>
			  	  <td height="27"></td>			  	 
			  	  <td height="27" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=strNomeAcao%></font></td>
		  	    </tr>
			  	<tr>
			  	  <td height="7"></td>
			  	  <td height="7" valign="middle" align="left"></td>
			  	  <td height="7" valign="middle" align="left"></td>
		  	    </tr>
				
				<tr> 
				  <td width="190" height="43"></td>
				  <td width="151" height="43" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Descrição do Corte:</b></font></td>
				  <td height="43" valign="middle" align="left" width="340">				  	
						<input type="hidden" name="txtCdCorte" value="<%=strCdCorte%>">					
						<input type="text" name="txtNomeCorte" maxlength="50" size="50" value="<%=strNomeCorte%>">	
							  
				  </td>
				</tr> 
								
				<tr> 
				  <td width="190" height="28"></td>
				  <td width="151" height="28" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data do Corte:</b></font></td>
				  <td height="28" valign="middle" align="left" width="340"> 
					<input type="text" name="txtDtCorte" maxlength="10" size="10" value="<%=strDtCorte%>">
				    <a href="javascript:show_calendar(true,'frmCadCorte.txtDtCorte','DD/MM/YYYY')"><img src="../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>			
				  </td>
				</tr>
								
				<tr> 
				  <td width="190" height="1"></td>
				  <td width="151" height="1" valign="middle" align="left"></td>
				  <td height="1" valign="middle" align="left" width="340"> </td>
				</tr>   
		  </table>
	</form>
	
<div id="aviso" style="position:absolute; width:356px; height:26px; z-index:1; left: 446px; top: 349px"></div>
</body>
	<%	
	db_banco.close
	set db_banco = nothing
	%>
</html>

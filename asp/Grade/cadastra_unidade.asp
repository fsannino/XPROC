<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))

strMostraOrgaoMenor = "Não"

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

if Request("selCorte") <> "" then
	intCdCorte = Request("selCorte")
elseif session("Corte") <> "" then
	intCdCorte = session("Corte")
end if

session("Corte") = intCdCorte

intCdUnidade = Request("selUnidade")	

'Response.write "selOrgaoMaior - " & trim(request("selOrgaoMaior")) & "<br>"

if trim(request("selOrgaoMaior")) <> "" then
	strOrgaoMaior = trim(request("selOrgaoMaior"))
	strOrgaoMaior = right((left(strOrgaoMaior,5)),3)	

	if(left(strOrgaoMaior,1)) = 0 then
		strOrgaoMaior = right(strOrgaoMaior,(len(strOrgaoMaior))-1)
	end if
else
	strOrgaoMaior = "000"
end if

'Response.write "strOrgaoMaior = " & strOrgaoMaior & "<br>"

if trim(request("selOrgaoMenor")) <> "" then
	strOrgaoMenor = trim(request("selOrgaoMenor"))
else
	strOrgaoMenor = 0
end if

strDiretoria  = request("selDiretoria")
strCT  = request("selCT")
strNomeUnidade = request("txtNomeUnidade")

'response.Write(strDiretoria)
'response.Write(strCT)

if request("txtOrgSel") <> "" then
   str_OrgSel = request("txtOrgSel")
else
   str_OrgSel = ""
end if

'intCdMultiplicador = trim(Request("selMultiplicador"))
'Response.write strAcao & "<br>"
'Response.write intCdMultiplicador & "<br>"
'Response.end

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
strSQL_CT = strSQL_CT & "WHERE CORT_CD_CORTE = " & intCdCorte
strSQL_CT = strSQL_CT & "ORDER BY CTRO_TX_NOME_CENTRO_TREINAMENTO"
'Response.WRITE  strSQL_CT & "<br><br>"
'Response.END
set rdsCT = db_banco.execute(strSQL_CT)

'******** ORGÃO MAIOR ******************
strSQL_OrgaoMaior = ""
strSQL_OrgaoMaior = "SELECT ORLO_CD_ORG_LOT, "
strSQL_OrgaoMaior = strSQL_OrgaoMaior & "ORLO_SG_ORG_LOT, ORLO_NM_ORG_LOT, ORLO_CD_STATUS "
strSQL_OrgaoMaior = strSQL_OrgaoMaior & "FROM GRADE_ORGAO_MAIOR "
strSQL_OrgaoMaior = strSQL_OrgaoMaior & "WHERE ORLO_CD_STATUS = 'A' "
strSQL_OrgaoMaior = strSQL_OrgaoMaior & "AND CORT_CD_CORTE = " & intCdCorte
strSQL_OrgaoMaior = strSQL_OrgaoMaior & " ORDER BY ORLO_SG_ORG_LOT"
'Response.WRITE  strSQL_OrgaoMaior & "<br><br>"
'Response.END
	
set rstOrgaoMaior = db_banco.execute(strSQL_OrgaoMaior)

if strAcao = "A" then		
		
	strSQLAltUnidade = ""
	strSQLAltUnidade = strSQLAltUnidade & "SELECT CTRO_CD_CENTRO_TREINAMENTO, "
	strSQLAltUnidade = strSQLAltUnidade & "ORLO_CD_ORG_LOT_DIR, UNID_TX_DESC_UNIDADE "
	strSQLAltUnidade = strSQLAltUnidade & "FROM GRADE_UNIDADE " 
	strSQLAltUnidade = strSQLAltUnidade & "WHERE UNID_CD_UNIDADE = " & intCdUnidade 
	strSQLAltUnidade = strSQLAltUnidade & " AND CORT_CD_CORTE = " & intCdCorte 
	'Response.write strSQLAltUnidade & "<br><br>"
	'Response.end
	
	Set rdsAltUnidade = db_banco.Execute(strSQLAltUnidade)			
	
	if not rdsAltUnidade.EOF then			
		
		if trim(Request("selCT")) <> "" then		    
			strCT = trim(Request("selCT"))
		else
			if not IsNull(rdsAltUnidade("CTRO_CD_CENTRO_TREINAMENTO"))	 then
				strCT = rdsAltUnidade("CTRO_CD_CENTRO_TREINAMENTO")	
			else
				strCT = 0
			end if	
		end if
		
		if trim(Request("selDiretoria")) <> "0" then
			strDiretoria = trim(Request("selDiretoria"))
		else
			strDiretoria = rdsAltUnidade("ORLO_CD_ORG_LOT_DIR")		
		end if
		
		if trim(Request("txtNomeUnidade")) <> "" then
			strNomeUnidade = trim(Request("txtNomeUnidade"))
		else
			strNomeUnidade 	= rdsAltUnidade("UNID_TX_DESC_UNIDADE")	
		end if
																		
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & "SELECT UNID_ORGAO.ORME_CD_ORG_MENOR "
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & "FROM GRADE_UNIDADE UNID, GRADE_UNIDADE_ORGAO_MENOR UNID_ORGAO "
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & "WHERE UNID.UNID_CD_UNIDADE = UNID_ORGAO.UNID_CD_UNIDADE "
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & "AND UNID.UNID_CD_UNIDADE = " & intCdUnidade 
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & " AND UNID.CORT_CD_CORTE = " & intCdCorte 
		strSQLUnidadeOrgaoMenor = strSQLUnidadeOrgaoMenor & " AND UNID_ORGAO.CORT_CD_CORTE = " & intCdCorte 
		'Response.write strSQLUnidadeOrgaoMenor
		'Response.end
	
		Set rdsAltUnidadeOrgaoMenor = db_banco.Execute(strSQLUnidadeOrgaoMenor)	
									
		if not rdsAltUnidadeOrgaoMenor.EOF then				
			strMostraOrgaoMenor = "Sim"				
			
			'Response.write "selOrgaoMaior 2 = " & request("selOrgaoMaior") & "<br>"
			'Response.write "strOrgaoMaior = " & strOrgaoMaior & "<br>"
			
			if trim(request("selOrgaoMaior")) = "" then					
				
				strOrgaoMaior = trim(rdsAltUnidadeOrgaoMenor("ORME_CD_ORG_MENOR"))
				
				'Response.write len(strOrgaoMaior) & "<br>"
				
				'if len(strOrgaoMaior) = 15 then
					strOrgaoMaior = right(left(strOrgaoMaior,5),3)	
							
					if left(strOrgaoMaior,1) = 0 then
						strOrgaoMaior = right(strOrgaoMaior,(len(strOrgaoMaior))-1)
					end if			
				'elseif len(strOrgaoMaior) = 14 then
					'strOrgaoMaior = right(left(strOrgaoMaior,4),3)								
				'end if		
			end if			
			'Response.write "strOrgaoMaior = " & strOrgaoMaior & "<br>"
			'Response.end							
		end if								
	end if
	
	rdsAltUnidade.close
	set rdsAltUnidade = nothing	
end if

'Response.write strOrgaoMaior
'Response.end

'******** ORGÃO MENOR ******************
strSQL_OrgaoMenor = ""
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "SELECT ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "FROM GRADE_ORGAO_MENOR "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "WHERE (ORLO_CD_ORG_LOT = " & strOrgaoMaior & ") "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "AND (ORME_CD_STATUS = 'A') "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "AND SUBSTRING(ORME_CD_ORG_MENOR,3,3) = '" & right("000"& strOrgaoMaior,3) & "' "
strSQL_OrgaoMenor = strSQL_OrgaoMenor & "AND CORT_CD_CORTE = " & intCdCorte 
'strSQL_OrgaoMenor = strSQL_OrgaoMenor & " ORDER BY ORME_NM_ORG_MENOR"
'Response.WRITE  strSQL_OrgaoMenor & "<br><br>"
'Response.END

set rstOrgaoMenor = db_banco.execute(strSQL_OrgaoMenor)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript">
					
			function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
			  var obj = MM_findObj(objName);
			  var obj2 = MM_findObj(theValue);
			  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
			  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
			}
			
			function MM_swapImgRestore() { //v3.0
			  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
			}
			
			function MM_preloadImages() { //v3.0
			  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
				var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
				if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
			}
			
			function MM_swapImage() { //v3.0
			  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
			   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
			}
			
			function MM_findObj(n, d) { //v4.01
			  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
			  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && d.getElementById) x=d.getElementById(n); return x;
			}		
					
			function Confirma()
			{			
				if(document.frmCadMultiplicador.selCorte.selectedIndex == 0)
				{
					alert("É obrigatória a seleção de um CORTE!");
					document.frmCadMultiplicador.selCorte.focus();
					return;
				}						
					
				if(document.frmCadMultiplicador.selDiretoria.selectedIndex == 0)
				{
					alert("É obrigatória a seleção de um DIRETORIA!");
					document.frmCadMultiplicador.selDiretoria.focus();
					return;
				}						
				
				if(document.frmCadMultiplicador.selCT.selectedIndex == 0)
				{
					alert("É obrigatória a seleção de um CENTRO DE TREINAMENTO!");
					document.frmCadMultiplicador.selCT.focus();
					return;
				}			
				
				if(document.frmCadMultiplicador.txtNomeUnidade.value == "")
				{
					alert("É necessário o preenchimento do campo NOME DA UNIDADE!");
					document.frmCadMultiplicador.txtNomeUnidade.focus();
					return;
				}					
				
				if(document.frmCadMultiplicador.listResultOrgao.options.length == 0)
				{					
					alert("É obrigatória a seleção de pelo menos um ORGÃO MENOR!");
					document.frmCadMultiplicador.listResultOrgao.focus();
					return;					
				}	
				
				if(document.frmCadMultiplicador.listResultOrgao.options.length == 1)
				{					
					if (document.frmCadMultiplicador.listResultOrgao(0).value == 0)
					{
						alert("É obrigatória a seleção de pelo menos um ORGÃO MENOR!");
						document.frmCadMultiplicador.listResultOrgao.focus();
						return;
					}
				}			
														
				//*** Monta uma string com os CURSOS Selecionados, separados por vírgula
				//carrega_txt(document.frmCadMultiplicador.selCurso_Selecionado)										
														
				document.frmCadMultiplicador.action="grava_unidade.asp";				
				document.frmCadMultiplicador.submit();			
			}		
			
			function submet_pagina(strvalor,strTipo)
			{					
				strCorte = document.frmCadMultiplicador.selCorte.value;
				strDiretoria = document.frmCadMultiplicador.selDiretoria.value;
				strCT = document.frmCadMultiplicador.selCT.value;
				strNomeUnidade = document.frmCadMultiplicador.txtNomeUnidade.value;
				
				//alert(strCorte + ' - ' + strDiretoria  + ' - ' + strCT  + ' - ' + strNomeUnidade);
				
				if (strTipo == 'OrgaoMaior')
				{
					document.frmCadMultiplicador.listResultOrgao.options.length = 0;		
					document.frmCadMultiplicador.txtOrgSel.value = '';									
				}														
						
				document.frmCadMultiplicador.action="cadastra_unidade.asp?selCorte="+strCorte+"&selDiretoria="+strDiretoria+"&selCT="+strCT+"&txtNomeUnidade="+strNomeUnidade;
				document.frmCadMultiplicador.submit();
			}
	
			function carrega_txt(fbox) 
			{
				document.frmCadMultiplicador.txtOrgSel.value = '';
				for(var i=0; i<fbox.options.length; i++) 
				{
					if (i == 0)
					{
						document.frmCadMultiplicador.txtOrgSel.value = fbox.options[i].value;
					}
					else
					{					
						document.frmCadMultiplicador.txtOrgSel.value = document.frmCadMultiplicador.txtOrgSel.value + "," + fbox.options[i].value;
					}
				}
			}	
			
			function Apaga_Item()
			{
				var f = document.frmCadMultiplicador.listResultOrgao.options.length;
				var items = '';
				
				for(var i = 0; i < f; i++)
				{
					if (document.frmCadMultiplicador.listResultOrgao.options[i].selected)
					{
					items = items + ';' + i
					}
				}
				
				items=items + ';';
				var t = document.frmCadMultiplicador.listResultOrgao.options.length;
				var f = -1;
				
				for(var d = 0; d < t + 1; d++)
				{
					var s = ';'+d+';';
					if(items.search(s)!=-1)
					{
						if(f==-1)
						{
						document.frmCadMultiplicador.listResultOrgao.options[d] = null;
						f=d;
						}
						else
						{
						document.frmCadMultiplicador.listResultOrgao.options[f] = null;
						}		
					}
				}
				carrega_txt(document.frmCadMultiplicador.listResultOrgao);		
			}				
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
	
		<script language="javascript" src="troca_lista_sem_retirar.js"></script>
	
		<form method="POST" name="frmCadMultiplicador">					
									
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Unidade - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="811" height="204">			  	
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
					 <td height="31" colspan="1"></td>
					 <td width="238" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
					 </td>
					 <td width="385" colspan="2" valign="middle">				
					   
					 <%
					 if strAcao = "A" then
					 
					 	strSQLCorte = ""
						strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
						'Response.write strSQLCorte
						'Response.end
						set rsCorte = db_banco.Execute(strSQLCorte)		
						
						strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")
					 	%>						
						<input type="hidden" name="selCorte" value="<%=cint(Session("Corte"))%>">	
						<font face="Verdana" size="2" color="#330099"><%=Ucase(strNomeCorte)%></font>						
					 	<%
					 else										 	
						strSQLCorte = ""
						strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						strSQLCorte = strSQLCorte & "ORDER BY CORT_DT_DATA_CORTE DESC"
						'Response.write strSQLCorte
						'Response.end
						set rsCorte = db_banco.Execute(strSQLCorte)												   
					 	%>					 
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
						<%						 
					end if	
					
					rsCorte.close
					set rsCorte = nothing		
					%>				   	   
					</td>
				  </tr>
			
				<tr>
				  <td height="26"></td>
				  <td height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>
				  <td height="26" valign="middle" align="left">				 
				  <select size="1" name="selDiretoria">
					<option value="0">== Selecione a Diretoria ==</option>
					<%
						do until rdsDiretoria.eof = true
							  if cint(strdiretoria) = cint(rdsDiretoria("ORLO_CD_ORG_LOT")) then%>
								<option value="<%=rdsDiretoria("ORLO_CD_ORG_LOT")%>" selected><%=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
							<%else%>
								<option value="<%=rdsDiretoria("ORLO_CD_ORG_LOT")%>"><%=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
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
				  <td width="174" height="26"></td>
				  <td width="238" height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:</b></font></td>
				  <td height="26" valign="middle" align="left" width="385">
					<select size="1" name="selCT">
					   <option value="0">== Selecione um Centro de Treinamento ==</option>
						<%					
						do until rdsCT.eof = true
							  if cint(strCT) = cint(rdsCT("CTRO_CD_CENTRO_TREINAMENTO")) then
							  %>
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
				  <td width="174" height="34"></td>
				  <td width="238" height="34" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome da Unidade:</b></font></td>
				  <td height="34" valign="middle" align="left" width="385">				  	
					<!--<input type="hidden" name="txtCdMultiplicador" value="<%'=intCdMultiplicador%>">	-->		
			 	  	<input type="text" name="txtNomeUnidade" maxlength="50" size="50" value="<%=strNomeUnidade%>">				  
				  </td>
				</tr>								
		  </table>		  
		  
		  
	  <table width="96%" height="163" border="0" cellspacing="2">
		<tr>
		  <td width="18%" height="1"></td>
		  <td height="1" colspan="3"></td>
		</tr>
		<tr>
		  <td height="24">&nbsp;</td>
		  <td width="37%" height="24"><b><font face="System" size="2" color="#330099">Unidade - &Oacute;rg&atilde;o maior:</font></b></td>
		  <td width="36%" rowspan="6" valign="bottom">
			<table width="100%"  border="0" cellspacing="0" cellpadding="1">
			  <tr>
				<td width="70%">
				  <font color="#000080" face="System" size="2">
					<select name="listResultOrgao" size="10" multiple>
					  	<%
						if strMostraOrgaoMenor = "Sim" then			
							
							contReg = 0							
							do while not rdsAltUnidadeOrgaoMenor.EOF
							
								strSQLNomeOrgaoMenor = ""
								strSQLNomeOrgaoMenor = strSQLNomeOrgaoMenor & "SELECT ORME_SG_ORG_MENOR "
								strSQLNomeOrgaoMenor = strSQLNomeOrgaoMenor & "FROM GRADE_ORGAO_MENOR "
								strSQLNomeOrgaoMenor = strSQLNomeOrgaoMenor & "WHERE ORME_CD_ORG_MENOR = '" & trim(rdsAltUnidadeOrgaoMenor("ORME_CD_ORG_MENOR")) & "'"
								strSQLNomeOrgaoMenor = strSQLNomeOrgaoMenor & " AND CORT_CD_CORTE = " & intCdCorte
								
								set rstNomeOrgao = db_banco.execute(strSQLNomeOrgaoMenor)
								
								if not rstNomeOrgao.eof then																	
									if trim(request("selOrgaoMaior")) = "" then										
										strNomeOrgao = trim(rstNomeOrgao("ORME_SG_ORG_MENOR"))
										%>
										<option value="<%=trim(rdsAltUnidadeOrgaoMenor("ORME_CD_ORG_MENOR"))%>"><%=strNomeOrgao%></option>
										<%										
										'*** CARREGA A VARIÁVEL COM OS VALORES CADASTRADOS NA BASE
										if contReg = 0 then
											str_OrgSel = trim(rdsAltUnidadeOrgaoMenor("ORME_CD_ORG_MENOR"))
										else
											str_OrgSel = str_OrgSel & "," & trim(rdsAltUnidadeOrgaoMenor("ORME_CD_ORG_MENOR"))
										end if											
									end if													
								end if										
								
								rdsAltUnidadeOrgaoMenor.movenext
								contReg = contReg + 1
							loop		
						end if					
						%>
				  </select>
				</font>
				</td>
				<td width="30%">
				<table width="12%" height="70" border="0" cellpadding="1" cellspacing="5">
				  <tr>
					<td><a href="#"><img src="../../imagens/delete_98.gif" alt="Excluir &Iacute;tem Selecionado" width="24" height="24" border="0" onClick="Apaga_Item()"></a></td>
				  </tr>
				</table>           
				</td>
			  </tr>
		  </table>      </td>
		  <td width="9%" height="24">&nbsp;</td>
		</tr>
		<tr>
		  <td width="18%" height="54"><div align="right"><b></b></div></td>
		  <td height="54">  	    <table width="99%"  border="0" cellspacing="0" cellpadding="1">
				  <tr>
				  	<td width="72%">
						<select name="selOrgaoMaior" size="1" onChange="submet_pagina(this.value,'OrgaoMaior')">
						  <option value="000">== Todas ==</option>
						  <%					  
						  do until rstOrgaoMaior.eof=true
							if cint(strOrgaoMaior) = cint(rstOrgaoMaior("ORLO_CD_ORG_LOT"))then
							  %>
							  <!--<option selected value="<%'="00" & right(("000" & rstOrgaoMaior("ORLO_CD_ORG_LOT")),3) & right(("000" & rstOrgaoMaior("ORLO_NR_ORDEM")),2)%>"><%'=rstOrgaoMaior("ORLO_SG_ORG_LOT")%></option>-->
							  	<option selected value="<%=rstOrgaoMaior("ORLO_CD_ORG_LOT")%>"><%=rstOrgaoMaior("ORLO_SG_ORG_LOT")%></option>
							  <%else%>
							  <!--<option value="<%'="00" & right(("000" & rstOrgaoMaior("ORLO_CD_ORG_LOT")),3) & right(("000" & rstOrgaoMaior("ORLO_NR_ORDEM")),2)%>"><%'=rstOrgaoMaior("ORLO_SG_ORG_LOT")%></option>-->
							  <option value="<%=rstOrgaoMaior("ORLO_CD_ORG_LOT")%>"><%=rstOrgaoMaior("ORLO_SG_ORG_LOT")%></option>
							  <%
							end if
							rstOrgaoMaior.movenext
						  looP
						  %>
						</select>
					</td>				
				  </tr>
		  </table>	  
		  </td>
		  <td height="54">&nbsp;</td>
		</tr>
		<tr>
		  <td height="13">&nbsp;</td>
		  <td height="13"><font color="#330099" face="System" size="2"><b>Gerência - &Oacute;rg&atilde;o menor:</b></font></td>
		  <td height="13">&nbsp;</td>
		</tr>
		<tr>
		  <td width="18%" height="13"><div align="right"><font color="#000080" face="System" size="2"><b> </b></font></div></td>
		  <td height="13">	  	<table width="100%"  border="0" cellspacing="0" cellpadding="1">
			  <tr>
				 <td width="72%">
					<select size="10" name="selOrgaoMenor" multiple>
					  <option value="0">== Todas ==</option>
						<%					
						do until rstOrgaoMenor.eof=true
							if trim(strOrgaoMenor) = trim(rstOrgaoMenor("ORME_CD_ORG_MENOR")) then 'left((rstOrgaoMenor("ORME_CD_ORG_MENOR")),10)
								%>
								<option selected value="<%=trim(rstOrgaoMenor("ORME_CD_ORG_MENOR"))%>"><%=rstOrgaoMenor("ORME_SG_ORG_MENOR")%></option>
								<%
							else
								%>
								<option value="<%=trim(rstOrgaoMenor("ORME_CD_ORG_MENOR"))%>"><%=rstOrgaoMenor("ORME_SG_ORG_MENOR")%></option>
								<%
							end if
							rstOrgaoMenor.movenext
						looP
						%>
					</select>			</td>
				<td width="28%" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1612','','../../imagens/continua_F02.gif',1)" onClick="move(document.frmCadMultiplicador.selOrgaoMenor,document.frmCadMultiplicador.listResultOrgao,0);carrega_txt(document.frmCadMultiplicador.listResultOrgao);"><img src="../../imagens/continua_F01.gif" alt="Incluir &Iacute;tem" name="Image1612" width="25" height="24" border="0" id="Image161"></a><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1611112','','../../imagens/continua_F02.gif',1)" onClick="move_tudo(document.frmCadMultiplicador.selOrgaoMenor,document.frmCadMultiplicador.listResultOrgao,0);carrega_txt(document.frmCadMultiplicador.listResultOrgao);"><img src="../../imagens/seta_dupla_direita.gif" alt="Incluir todos os Ítem" name="Image1611112" width="24" height="24" border="0" id="Image161"></a></td>
			  </tr>
			</table>
		  </td>
		  <td height="13">&nbsp;</td>
		</tr>  
	  </table>
	  
	  <input type="hidden" name="pAcao" value="<%=strAcao%>"> 	
	  <input type="hidden" name="txtOrgSel" value="<%=str_OrgSel%>">	
	  <input type="hidden" name="selUnidade" value="<%=intCdUnidade%>"> 
			
  	</form>
	</body>
	<%		
	rdsAltUnidadeOrgaoMenor.close
	set rdsAltUnidadeOrgaoMenor = nothing		
		
	rstOrgaoMenor.close
	set rstOrgaoMenor = nothing
		
	rstOrgaoMaior.close
	set rstOrgaoMaior = nothing
	
	db_banco.close
	set db_banco = nothing
	%>

</html>

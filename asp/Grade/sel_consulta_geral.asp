<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
Response.Expires = 0
Session.LCID = 1046
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		
		<style type="text/css">
			<!--
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
			-->
		</style>
				
	</head>

	<script language="javascript">
	
		function Confirma()
		{		
			//var intOndaAnt	= 0;
			//var intEaD		= 0;
			//var intDescentr = 0;
			//var intInLoco	= 0;
			
			//if (document.frm1.selDiretoria.selectedIndex == 0)
			//{
				//alert('Para consultar é necessária a escolha de uma Diretoria!');
				//document.form1.selDiretoria.focus();
				//return;
			//}
			
			////if (document.frm1.chkOndaAntecipada.checked == true)
			////{
				////intOndaAnt = 1;
			////}
		
			//if (document.frm1.chkEaD.checked == true)
			//{
				//intEaD = 1;
			//}
			
			//if (document.frm1.chkDescentralizado.checked == true)
			//{
				//intDescentr = 1;
			//}
			
			//if (document.frm1.chkInLoco.checked == true)
			//{
				//intInLoco = 1;
			//}
			
			//alert('intOndaAnt - ' + intOndaAnt);
			//alert('intEaD - ' + intEaD);
			//alert('intDescentr - ' + intDescentr);
			//alert('intInLoco - ' + intInLoco);
												
			//document.frm1.action = "gera_consulta_turma.asp?pOndaAnt="+intOndaAnt+"&pEaD="+intEaD+"&pDescentr="+intDescentr+"&pInLoco="+intInLoco;									
			//document.frm1.action = "gera_consulta_turma.asp?pEaD="+intEaD+"&pDescentr="+intDescentr+"&pInLoco="+intInLoco;
			document.frm1.action = "gera_consulta_turma.asp";
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
					<td width="26"><!--<a href="javascript:Confirma();"><img border="0" src="../Funcao/confirma_f02.gif"></a>--></td>
				  <td width="50"><!--<font color="#330099" face="Verdana" size="2"><b>Enviar</b></font>--></td>
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
			<div align="center"><font face="Verdana" color="#330099" size="3"><b>Grade de Treinamento<b></font></div>
		  </td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		</tr>
	  </table>
		  
	  <table border="0" width="849" height="92">						
		<tr> 
		  <td width="222" height="57"></td>
		  <td width="224" height="57" valign="middle" align="left"><font face="Verdana" color="#330099" size=""><b>Relatórios:</b></font></td>
		  <td height="57" valign="middle" align="left" width="389"></td>
		</tr> 
		<tr> 
		  <td width="222" height="21"></td>
		  <td width="224" height="21" valign="middle" align="left">
		  	<font face="Verdana" color="#330099" size="2">-</font>&nbsp;
		  	<a href="consulta_geral.asp?pNumRel=1&pTituloRel=Demanda x Oferta EAD" id="Link">Demanda x Oferta EAD</a></td>
		  <td height="21" valign="middle" align="left" width="389"></td>
		</tr>    
		
		<tr> 
		  <td width="222" height="21"></td>
		  <td width="224" height="21" valign="middle" align="left">
		  	<font face="Verdana" color="#330099" size="2">-</font>&nbsp;
		  	<a href="consulta_geral.asp?pNumRel=2&pTituloRel=Demanda Curso" id="Link">Demanda Curso</a></td>
		  <td height="21" valign="middle" align="left" width="389"></td>
		</tr>      
	  </table>
</form>
	</body>
</html>

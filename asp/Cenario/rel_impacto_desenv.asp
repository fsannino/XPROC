<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = 0

str_Assunto=0
str_Assunto=request("selAssunto")

str_Opc = Request("txtOpc")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

SQL_Desenvolvimento = ""
SQL_Desenvolvimento = SQL_Desenvolvimento & "SELECT DESE_CD_DESENVOLVIMENTO "
SQL_Desenvolvimento = SQL_Desenvolvimento & "FROM " & Session("PREFIXO") & "DESENVOLVIMENTO "
SQL_Desenvolvimento = SQL_Desenvolvimento & "ORDER BY DESE_CD_DESENVOLVIMENTO"
set rs_desenvolvimento = conn_db.execute(SQL_Desenvolvimento)
%>
<html>
<head>
	<STYLE type=text/css>
		BODY {
		SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
	</STYLE>
	<title></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript">			
		function Confirma() 
		{
			var intSelDesenv = document.frm1.selDesenvolvimento.selectedIndex;
			var strTxtDesenv = document.frm1.txtDesenvolvimento.value;
			
			if (intSelDesenv == 0 && strTxtDesenv == '')
			{
			  alert("Você deve selecionar ou digitar um Desenvolvimento!");				
			  document.frm1.selDesenvolvimento.focus();
			  return;								 
			}
			
			if (intSelDesenv != 0 && strTxtDesenv != '')
			{
			  alert("Você deve escolher entre selecionar uma opção ou digitar um Desenvolvimento!");				
			  document.frm1.selDesenvolvimento.focus();
			  return;								 
			}			
			document.frm1.submit();
			
		}
		
		function Limpa()
		{
			document.frm1.reset();
		}
		
		function MM_swapImgRestore() { //v3.0
		  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
		}
		
		function MM_preloadImages() { //v3.0
		  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
			var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
			if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
		}
		
		function MM_findObj(n, d) { //v4.0
		  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
			d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
		  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
		  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
		  if(!x && document.getElementById) x=document.getElementById(n); return x;
		}
		
		function MM_swapImage() { //v3.0
		  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
		   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
		}
		
		function VerificaPadraoDesenv(strvalor,strNome)	
		{	
			if (document.forms[0].txtDesenvolvimento.value != '')
			{									
				var strvalor = new String(strvalor);						
				var int_tamanho = strvalor.length;											
				
				//*** Referente as Letras																
				if (int_tamanho <= 2)
				{		
					var strLetra1 = '';	
					var strLetra2 = '';	
										
					if (int_tamanho==1)strLetra1 = strvalor.substring(0,1);	
					if (int_tamanho==2)strLetra2 = strvalor.substring(1,2);				
					
					if (strLetra1 != '')
					{
						if (isNaN(strLetra1)==false)
						{					
							alert("Os dois primeiro caracteres do código do desenvolvimento devem ser letras!");
							document.forms[0].txtDesenvolvimento.value = '';	
							document.forms[0].txtDesenvolvimento.focus();
							return;				
						}
					}
					
					if (strLetra2 != '')
					{
						if (isNaN(strLetra2)==false)
						{					
							alert("Os dois primeiro caracteres do código do desenvolvimento devem ser letras!");
							document.forms[0].txtDesenvolvimento.value = '';	
							document.forms[0].txtDesenvolvimento.focus();
							return;				
						}
					}
				}
							
				//*** Referente aos Números					
				if ((int_tamanho > 2) && (int_tamanho < 7))
				{				
					var strNumero = strvalor.substring(2,6);							
					if (isNaN(strNumero))
					{					
						alert("Os quatro últimos caracteres do código do desenvolvimento devem ser números!");
						document.forms[0].txtDesenvolvimento.value = '';	
						document.forms[0].txtDesenvolvimento.focus();
						return;				
					}
				}	
			}			
		}	
		
		function VerifiCacaretersEspeciais(strvalor,strobjnome)
		{			
			var vetEspeciais = new Array();			
			var strvalor = new String(strvalor);		
						
			var i, j;
			vetEspeciais[0] = "&";
			vetEspeciais[1] = "'";
			vetEspeciais[2] = '"'
			vetEspeciais[3] = '>';
			vetEspeciais[4] = '<';	
			vetEspeciais[5] = '.';	
			vetEspeciais[6] = ',';	
			vetEspeciais[7] = ' ';		
			vetEspeciais[8] = ':';	
						
			i=0;
			j=0;
						
			for (i=0; i<=strvalor.length-1; i++)
			{			
				for (j=0; j<=vetEspeciais.length-1; j++)
				{					
					if (strvalor.charAt(i) == vetEspeciais[j])
					{
						alert ('O caracter ' + strvalor.charAt(i) + ' não pode ser utilizado no texto.');
						
						if (strobjnome=='txtDesenvolvimento')
						{
							document.forms[0].txtDesenvolvimento.value = strvalor.substr(0,i);
						}
						break;
					}
				}
			}		
		}
	</script>
</head>

<SCRIPT LANGUAGE="JavaScript">
	function addbookmark()
	{
		bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
		bookmarktitle="Sinergia - Cadastro"
		if (document.all)
		window.external.AddFavorite(bookmarkurl,bookmarktitle)
	}	
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif')">
<form name="frm1" method="post" action="gera_rel_impacto_desenv.asp">
  <input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma();"><img src="confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="96%" border="0" cellpadding="0" cellspacing="5" name="tblSubProcesso" height="146">
    <tr> 
      <td width="17%" height="75">&nbsp;</td>
      <td width="62%" height="75"> 
        <input type="hidden" name="txtOpc" value="1">
        <p align="center"><font face="Verdana" color="#330099" size="3">Relatório de Cenários Impactados por Desenvolvimentos</font>
        <p align="center" style="margin-top: 0; margin-bottom: 0"> 
      </td>
      <td width="7%" height="75">       
      </td>
      <td width="4%" height="75">     
      </td>
    </tr>	
    <tr>
      <td>
        <p align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Desenvolvimento:</font></b> </td>
      <td>
        <select size="1" name="selDesenvolvimento">
          <option value="0">Selecione um Desenvolvimento</option>
            <%
		    do until rs_desenvolvimento.eof = true%>
			  <option value=<%=rs_desenvolvimento("DESE_CD_DESENVOLVIMENTO")%>><%=rs_desenvolvimento("DESE_CD_DESENVOLVIMENTO")%></option>
			  <%
			  rs_desenvolvimento.movenext
			loop
			%>
        </select>
		<%
		rs_desenvolvimento.close
		set rs_desenvolvimento = nothing
		%>
        </td>
      <td height="1">&nbsp;</td>
      <td height="1">&nbsp;</td>
    </tr>		
    <tr>
      <td height="21" colspan="4">	  
		  <table width="97%" height="46" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso">
			<tr> 		
			  <td height="21" width="3%"></td>	  
			  <td height="21" colspan="3"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">&nbsp;Se preferir, digite o código do Desenvolvimento abaixo</font></b></td>			  
			</tr>
			<tr> 
			  <td width="3%" height="25"></td>
			  <td width="15%" height="21"></td>	
			  <td width="49%" height="25">
			  	<input type="text" name="txtDesenvolvimento" size="20" maxlength="6" onKeyUp="javascript:VerificaPadraoDesenv(this.value,this.name);VerifiCacaretersEspeciais(this.value,this.name);"> 
				<font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1">( Ex.: MM0100 )</font> 
			  </td>
			  <td width="10%" height="25">&nbsp;</td>
			  <td width="23%" height="25">&nbsp;</td>
			</tr>
		  </table>	  
	  </td>      
    </tr>    
  </table>
</form>
<p>&nbsp;</p>
</body>
</html>

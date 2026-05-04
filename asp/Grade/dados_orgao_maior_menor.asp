<%

'set db_banco = Server.CreateObject("AdoDB.Connection")
'db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
'db_banco.open Session("Conn_String_Cogest_Gravacao")
'db_banco.CursorLocation = 3

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

int_Tot_Registro = 0

strNumRel				= request("pNumRel")
strTituloRel			= request("pTituloRel")

'Request("selCorte"  & "<p>")
'Request("selDiretoria"  & "<p>")
'Request("selOrgaoMaior" & "<p>")

if trim(Request("selCorte")) <> "0" and trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))
end if 

if trim(Request("selDiretoria")) <> "0" then
	int_Cd_Diretoria = trim(Request("selDiretoria"))
else
	int_Cd_Diretoria = "0"
end if

if trim(Request("selOrgaoMaior")) <> "0" then
	int_Cd_Org_Maior = trim(Request("selOrgaoMaior"))
else
	int_Cd_Org_Maior = "0"
end if

%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
        <style type="text/css">
<!--
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; font-size: 12px; }
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
-->
        </style>
<script>
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
</script>			
</head>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">		
		<%		
		if request("excel") <> 1 then
		%>		
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
					<td width="26"></td>
				  <td width="50"></td>
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					<td width="27"></td>  
					<td width="50"></td>
				  <td width="28"></td>
				  <td width="26">&nbsp;</td>
				  <td width="159"></td>
				</tr>
			  </table>
			</td>
		  </tr>
		</table>	
		
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="561"></td>	
					<td width="237">
						<div align="center">	
							 <a href="dados_orgao_maior_menor.asp?excel=1&amp;selDiretoria=<%=int_Cd_Diretoria%>&amp;selOrgaoMaior=<%=int_Cd_Org_Maior%>&amp;pTituloRel=<%=strTituloRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
						</div>
					</td>						
				    <td width="190"><img src="../../Flash/preloader.gif" name="loader" width="190" height="50" id="loader"></td>
				</tr>
			</table>
		<%end if%>		
<table cellspacing="0" cellpadding="0" border="0" width="100%">
  <tr>
    <td height="10"> </td>
  </tr>
  <tr>
    <td>
      <div align="center"><font face="Verdana" color="#330099" size="3"><b>Relat&oacute;rio de <%=strTituloRel%> - Grade de Treinamento</b></font></div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>		
<%
str_Sql = " SELECT "
str_Sql = str_Sql & " dbo.GRADE_ORGAO_MAIOR.CORT_CD_CORTE"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MAIOR.ORLO_CD_ORG_LOT"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MAIOR.ORLO_SG_ORG_LOT"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MAIOR.ORLO_CD_GABINETE"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MAIOR.ORLO_CD_STATUS"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MENOR.CORT_CD_CORTE AS Expr1"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MENOR.ORME_CD_ORG_MENOR"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MENOR.ORME_SG_ORG_MENOR"
str_Sql = str_Sql & " , dbo.GRADE_ORGAO_MENOR.ORME_CD_STATUS"
str_Sql = str_Sql & " FROM dbo.GRADE_ORGAO_MAIOR "
str_Sql = str_Sql & " INNER JOIN dbo.GRADE_ORGAO_MENOR ON dbo.GRADE_ORGAO_MAIOR.CORT_CD_CORTE = dbo.GRADE_ORGAO_MENOR.CORT_CD_CORTE AND" 
str_Sql = str_Sql & " dbo.GRADE_ORGAO_MAIOR.ORLO_CD_ORG_LOT = dbo.GRADE_ORGAO_MENOR.ORLO_CD_ORG_LOT"
str_Sql = str_Sql & " WHERE dbo.GRADE_ORGAO_MAIOR.CORT_CD_CORTE = " & Session("Corte")
str_Sql = str_Sql & " AND dbo.GRADE_ORGAO_MENOR.CORT_CD_CORTE = " & Session("Corte")
if int_Cd_Diretoria <> "0" then
	str_Sql = str_Sql & " AND dbo.GRADE_ORGAO_MAIOR.ORLO_CD_GABINETE = " & int_Cd_Diretoria
end if
if int_Cd_Org_Maior <> "0" then
	str_Sql = str_Sql & " AND dbo.GRADE_ORGAO_MAIOR.ORLO_CD_ORG_LOT = " & int_Cd_Org_Maior
end if
str_Sql = str_Sql & " AND dbo.GRADE_ORGAO_MAIOR.ORLO_CD_STATUS = 'A'"
str_Sql = str_Sql & " AND dbo.GRADE_ORGAO_MENOR.ORME_CD_STATUS = 'A'"
str_Sql = str_Sql & " order by dbo.GRADE_ORGAO_MAIOR.ORLO_CD_ORG_LOT, dbo.GRADE_ORGAO_MENOR.ORME_CD_ORG_MENOR "

'response.Write(str_Sql)

set rds_Orgao_Maior_Menor = db_banco.Execute(str_Sql)
If not rds_Orgao_Maior_Menor.Eof then
%>	
<table width="83%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#CCCCCC">
    <td><strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o Maior </font></strong></td>
    <td><strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o Menor </font></strong></td>
  </tr>
  <%
	  Do while 	not rds_Orgao_Maior_Menor.Eof
  
  		if str_Cor_Linha = "#FFFFFF" then 
		   str_Cor_Linha = "#F1F1F1"
		else
		   str_Cor_Linha = "#FFFFFF"
		end if				

  %>
  <tr bgcolor="<%=str_Cor_Linha%>">
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Orgao_Maior_Menor("ORLO_SG_ORG_LOT")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Orgao_Maior_Menor("ORME_SG_ORG_MENOR")%></font></td>
  </tr>
  <%	int_Tot_Registro = int_Tot_Registro + 1
  		rds_Orgao_Maior_Menor.movenext
	Loop
  %>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>	
<p align="center" class="style5">Total de registros = <%=int_Tot_Registro%></p>
	<% 
	end if
	if int_Tot_Registro = 0 then %>
	    <table width="76%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="21%">&nbsp;</td>
          </tr>
          <tr>
            <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">N&atilde;o encontrado registros</font></div></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
	<%	
		End if 
		%>	
</body>	
	<%	
	rds_Orgao_Maior_Menor.close
	set rds_Orgao_Maior_Menor = nothing
	db_banco.close
	set db_banco = nothing
	%>		
<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>	
</html>

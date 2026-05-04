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

strTituloRel			= request("pTituloRel")

int_Cd_Corte1 = request("selCorte1")
int_Cd_Corte2 = request("selCorte2")

int_Tot_Registro = 0

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
							 <a href="gera_consulta_turma.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;rdDescentralizado=<%=strDescentralizado%>&amp;rdEad=<%=strEaD%>&amp;rdInLoco=<%=strInLoco%>&amp;selCurso=<%=strCurso%>&amp;pTituloRel=<%=strTituloRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
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
		
		int_Tot_Registro = 0
		
		str_Sql = ""
		str_Sql = str_Sql & " SELECT "
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & ".UNID_TX_DESC_UNIDADE"
		str_Sql = str_Sql & " FROM dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & " LEFT OUTER JOIN"
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & " ON "
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & ".UNID_TX_DESC_UNIDADE = dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & ".UNID_TX_DESC_UNIDADE"
		str_Sql = str_Sql & " WHERE dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & ".UNID_TX_DESC_UNIDADE IS NULL"
		str_Sql = str_Sql & " order by dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & ".UNID_TX_DESC_UNIDADE "		
		set rds_Corte_1 = db_Banco.Execute(str_Sql)
		if not rds_Corte_1.Eof then
		%>
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="21%"><div align="right"><strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Corte : <%=int_Cd_Corte1%></font></strong></div></td>
    <td width="46%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
  <% 
  			Do while not rds_Corte_1.Eof
	%>
  <tr>
    <td>&nbsp;</td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Corte_1("UNID_TX_DESC_UNIDADE")%></font></td>
    <td>&nbsp;</td>
  </tr>
  
  <%			int_Tot_Registro = int_Tot_Registro + 1
  				rds_Corte_1.movenext
  			Loop
		end if	
  %>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
	
<p align="center" class="style5">Total de registros = <%=int_Tot_Registro%></p>
<% if int_Tot_Registro = 0 then %>
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
            <% End if %>
		<%
		
		int_Tot_Registro = 0
		
		str_Sql = ""
		str_Sql = str_Sql & " SELECT "
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & ".UNID_TX_DESC_UNIDADE"
		str_Sql = str_Sql & " FROM dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & " LEFT OUTER JOIN"
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & " ON "
		str_Sql = str_Sql & " dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & ".UNID_TX_DESC_UNIDADE = dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & ".UNID_TX_DESC_UNIDADE"
		str_Sql = str_Sql & " WHERE dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte1 & ".UNID_TX_DESC_UNIDADE IS NULL"
		str_Sql = str_Sql & " order by dbo.GRADE_UNIDADE_CORTE_" & int_Cd_Corte2 & ".UNID_TX_DESC_UNIDADE "
		set rds_Corte_2 = db_Banco.Execute(str_Sql)
		if not rds_Corte_2.Eof then
		%>
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="21%"><div align="right"><strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Corte : <%=int_Cd_Corte2%></font></strong></div></td>
    <td width="46%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
  <% 
  			Do while not rds_Corte_2.Eof
	%>
  <tr>
    <td>&nbsp;</td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Corte_2("UNID_TX_DESC_UNIDADE")%></font></td>
    <td>&nbsp;</td>
  </tr>			
  <%			int_Tot_Registro = int_Tot_Registro + 1
  				rds_Corte_2.movenext
  			Loop
		end if	
  %>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
	
<p align="center" class="style5">Totao de registros = <%=int_Tot_Registro%></p>
<% if int_Tot_Registro = 0 then %>
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
            <% End if %>		
</body>	
	<%
	rds_Corte_1.close
	set rds_Corte_1 = nothing
	rds_Corte_2.close
	set rds_Corte_2 = nothing
	db_banco.close
	set db_banco = nothing
	%>		
<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>	
</html>

<%
Response.Expires=0
Response.Buffer = True

on error resume next	
	set db_Cronograma = Server.CreateObject("ADODB.Connection")
	db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

if err.number <> 0 then		
	strMSG = "Ocorreu algum problema com o servidor!"
	Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pErroServidor=S"
end if	

str_Sql_Projetos = ""
str_Sql_Projetos = str_Sql_Projetos & " SELECT   "
str_Sql_Projetos = str_Sql_Projetos & " PROJ_ID"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_NAME"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_AUTHOR"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_COMPANY"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_INFO_CAL_NAME"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_SUBJECT"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_TITLE"
str_Sql_Projetos = str_Sql_Projetos & " FROM MSP_PROJECTS"
'response.Write(str_Sql_Projetos)
'response.End()
set rds_Projeto=db_Cronograma.execute(str_Sql_Projetos)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Lista de Projetos</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">	
	function impressao() 
	{		
		window.open('impressao.asp?par_PaginaPrint=relat_imp_lista_projeto.asp','jan1','toolbar=no, location=no, scrollbars=no, status=no, directories=no, resizable=no, menubar=no, fullscreen=no, height=50, width=250, top=200, left=260');
	}
	
	function exporta() 
	{			
		window.open('relat_imp_lista_projeto.asp?str_Tipo_Saida=Excel','jan1','toolbar=yes, location=no, scrollbars=yes, status=no, directories=no, resizable=yes, menubar=yes, fullscreen=no, height=400, width=500, status=no, top=100, left=160');
	}
</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../indexA_xpep.asp"><img src="../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" --> 
	 <table width="800" border="0" cellspacing="0" cellpadding="0">
		<tr> 
		  <td width="3%">&nbsp;</td>
		  <td width="88%">&nbsp;</td>
		  <td width="9%">&nbsp;</td>
		</tr>
		<tr> 
		  <td>&nbsp;</td>
		  <td class="subtitulo"><div align="center" class="campob">Rela&ccedil;&atilde;o de Projetos</font></td>
		  <td>&nbsp;</td>
		</tr>
		<tr> 
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
        <tr>	
		   <td colspan="1">&nbsp;</td>		
			<td colspan="2">
	  	  
    <table width="100%" border="1" cellspacing="0" bordercolor="#999999">
		<tr bgcolor="#639ACE" class="titcoltabela"> 
		  <td width="62"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Id Projeto</font></div></td>
		  <td width="162"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</font></div></td>
		  <td width="108"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Autor</font></div></td>
		  <td width="124"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Company</font></div></td>
		  <td width="92"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cal Name</font></div></td>
		  <td width="108"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Subject</font></div></td>
		  <td width="90"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Titulo</font></div></td>
		</tr>		
		<% 
		if not rds_Projeto.EOF then
			Do while not rds_Projeto.EOF 
				if str_Cor01 = "#EEEEEE" then
				   str_Cor01 = "#FFFFFF"
				else
				   str_Cor01 = "#EEEEEE"		
				end if		
				
				if rds_Projeto("PROJ_NAME") <> "" then
					strNomeProjeto = Ucase(rds_Projeto("PROJ_NAME"))
				else
					strNomeProjeto = "-"
				end if
				
				if rds_Projeto("PROJ_PROP_AUTHOR") <> "" then
					strAutorProjeto = Ucase(rds_Projeto("PROJ_PROP_AUTHOR"))
				else
					strAutorProjeto = "-"
				end if
				
				if rds_Projeto("PROJ_PROP_COMPANY") <> "" then
					strCompanhia = Ucase(rds_Projeto("PROJ_PROP_COMPANY"))
				else
					strCompanhia = "-"
				end if		
				
				if rds_Projeto("PROJ_INFO_CAL_NAME") <> "" then
					strCalName = Ucase(rds_Projeto("PROJ_INFO_CAL_NAME"))
				else
					strCalName = "-"
				end if	
									
				if rds_Projeto("PROJ_PROP_SUBJECT") <> "" then
					strSubject = Ucase(rds_Projeto("PROJ_PROP_SUBJECT"))
				else
					strSubject = "-"
				end if	
									
				if rds_Projeto("PROJ_PROP_TITLE") <> "" then
					strTitulo = Ucase(rds_Projeto("PROJ_PROP_TITLE"))
				else
					strTitulo = "-"
				end if	  						
				%>
				<tr bgcolor="<%=str_Cor01%>"> 
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Projeto("PROJ_ID")%></font></div></td>
				  <td><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strNomeProjeto%></font></div></td>
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strAutorProjeto%></font></div></td>
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strCompanhia%></font></div></td>
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strCalName%></font></div></td>
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strSubject%></font></div></td>
				  <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=strTitulo%></font></div></td>
				</tr>
				<%   
				 rds_Projeto.movenext
		   Loop
		
		else
			str_Msg = "N&atilde;o existem registros para esta condi&ccedil;&atilde;o."
		end if	
		
		rds_Projeto.Close 
		set rds_Projeto = nothing
		db_Cronograma.Close
		%>
	</table>
	
	</td>
	</tr>
	</table>
	
	<%
	if str_Msg <> "" then 
	%>
    <table width="800"  border="0" cellspacing="0" cellpadding="1">
	  <% For i=1 to 5 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>
      <tr>
        <td width="146">&nbsp;</td>
        <td width="634"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Msg%></font></div></td>
        <td width="207">&nbsp;</td>
      </tr>
	  <% For j=1 to 2 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"></div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"></div></td>
        <td>&nbsp;</td>
      </tr>
	  <% For j=1 to 3 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>	  
    </table>
	<% end if %>
	<table width="800"  border="0" cellspacing="0" cellpadding="1">
  <tr>
    <td width="155">&nbsp;</td>
    <td width="156"><div align="center"></div></td>
    <td width="122">&nbsp;</td>
    <td width="359">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="center"><a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" alt=":: Volta tela anterior" width="85" height="19" border="0"></a></div></td>
    <td><a href="#"><img src="../img/botao_imprimir.gif" alt=":: Imprime formato relatório" width="85" height="19" border="0" onclick="impressao();"></a></td>
    <td><a href="#"><img src="../img/botao_exportar_excel.gif" alt=":: Exporta formato Excel" width="85" height="19" border="0" onclick="exporta();"></a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>

<%
Response.Expires=0
Response.Buffer = True

if request("str_Tipo_Saida")="Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

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
<html>
<script language="JavaScript" type="text/JavaScript">
	function MM_reloadPage(init) {  //reloads the window if Nav4 resized
	  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
		document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
	  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
	}
	MM_reloadPage(true);
</script>
<head>	
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
	<link href="../../../css/biblioteca.css" rel="stylesheet" type="text/css">
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
	<title></title>	
</head>

<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	 <table width="800" border="0" cellspacing="0" cellpadding="0">
		<tr> 
		  <td width="1%">&nbsp;</td>
		  <td width="90%">&nbsp;</td>
		  <td width="9%">&nbsp;</td>
		</tr>
		<tr>				 
		  <td class="subtitulob" colspan="2"><div align="center" class="campob">Rela&ccedil;&atilde;o de Projetos</font></td>
		</tr>
		<tr> 			  
		  <td colspan="2">&nbsp;</td>
		</tr>
        <tr>		   	
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
</body>
<script language="javascript">
	function fechar()
		{
		window.top.close();	
		}	
		
	setTimeout('fechar()',1);
	window.top.frame2.focus();
	window.top.frame2.print();
</script>
</html>
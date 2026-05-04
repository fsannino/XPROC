<%
MEGA_PROCESSO = Request.Form("MEGA_PROCESSO")
CURSO = Request.Form("CURSO")
OMAIOR = Request.Form("OMAIOR")
OMENOR = Request.Form("OMENOR")
%>
<html>
<head>
<title>Sinergia Cursos</title>
<LINK href="../../style/estilo.css" type=text/css rel=styleSheet>
<SCRIPT>
	function fun_voltar() {
		top.parent.inf.location.href = 'progeven_petrobras.asp';
	}
</SCRIPT>
</head>
<body>
<%
	'Response.Write(CURSO)
	

   set conn=Server.CreateObject("ADODB.Connection")
   conn.Open Application("str_conn_cogest")
   set conn2=Server.CreateObject("ADODB.Connection")
   conn2.Open Application("str_conn")

   sQuery = "SELECT CURS_TX_NOME_CURSO "
   sQuery = sQuery & " FROM " & Application("owner") & "CURSO" 
   sQuery = sQuery & " WHERE CURS_CD_CURSO = '" & CURSO & "'"
	 
	 set obj_RS = Conn.Execute(sQuery)
   'NOME_CURSO = obj_RS("CURS_TX_NOME_CURSO")

   '-----   
   'PONTEIRO = 16
   'CARACTER = "0"

   'WHILE CARACTER = "0"
		'PONTEIRO = PONTEIRO - 1
		'CARACTER = MID(OMENOR, PONTEIRO, 1)  

   if NOT MID(OMENOR, 14, 2) = "00" then
   	PONTEIRO = 15
   else if NOT MID(OMENOR, 11, 3) = "000" then
   		PONTEIRO = 13
   	else 	if NOT MID(OMENOR, 8, 3) = "000" then
		   PONTEIRO = 10
   		else PONTEIRO = 7
   		end if
   	end if
   end if

   '-----

   sQuery = "SELECT		DISTINCT " & Application("owner") & "USUARIO_MAPEAMENTO.USMA_CD_USUARIO, " & _
			"			" & Application("owner") & "USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, " & _
			"			" & Application("owner") & "ORGAO_MENOR.ORME_SG_ORG_MENOR, CURSO.CURS_CD_CURSO AS COD_DISCIPLINA " & _
			"FROM		" & Application("owner") & "USUARIO_MAPEAMENTO INNER JOIN " & _
			"			" & Application("owner") & "FUNCAO_USUARIO ON " & _
			"			" & Application("owner") & "USUARIO_MAPEAMENTO.USMA_CD_USUARIO = " & Application("owner") & "FUNCAO_USUARIO.USMA_CD_USUARIO " & _
			"INNER JOIN " & Application("owner") & "ORGAO_MENOR ON " & _
			"			" & Application("owner") & "USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = " & Application("owner") & "ORGAO_MENOR.ORME_CD_ORG_MENOR " & _
			"LEFT OUTER JOIN " & Application("owner") & "CURSO_FUNCAO ON " & _
			"			" & Application("owner") & "FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = " & Application("owner") & "CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO " & _
			"LEFT OUTER JOIN " & Application("owner") & "CURSO ON " & _
			"			" & Application("owner") & "CURSO_FUNCAO.CURS_CD_CURSO = " & Application("owner") & "CURSO.CURS_CD_CURSO " & _

			"WHERE	ORME_CD_STATUS = 'A' " &_
                        " AND (" & Application("owner") & "FUNCAO_USUARIO.FUUS_IN_PRIORITARIO='1' " & ") " &_			
			" AND (" & Application("owner") & "CURSO.CURS_CD_CURSO='" & Request.Form("CURSO") & "') "
			
			if Request.Form("OMAIOR") <> "0" AND Request.Form("OMENOR") <> "0" then
							sQuery = sQuery & " AND SUBSTRING(" & Application("owner") & "ORGAO_MENOR.ORME_CD_ORG_MENOR, 1, " & CSTR(PONTEIRO) & ") = '" & MID(OMENOR, 1, PONTEIRO) & "' "		
			end if
			
		if MEGA_PROCESSO <> "0" then
			sQuery = sQuery & " AND (" & Application("owner") & "CURSO.MEPR_CD_MEGA_PROCESSO = " & MEGA_PROCESSO & ") "
		end if

			sQuery = sQuery & " ORDER BY	3, 2"

	'Response.Write sQuery
				
   set RS_Proj = Conn.Execute(sQuery)
      	
%>

 <TABLE width="760" align=center>
<TR>
<TD WIDTH="660"><font class="Titulo"><% =CURSO %> - RELAÇÃO DE INDICADOS PRIORITÁRIOS<br><BR></font></TD>
<TD ALIGN="RIGHT">      <div align="right"><a href="javascript:self.print()"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Imprimir</font></a></div></TD>
</TR>


<% If not RS_Proj.eof Then  %>

 <% Totaluno = 0 %>
 <% Totaluno_aprov = 0 %>
 <TR>
 <TD>
 <TABLE border=1>
     	<TR>

		<% if Request.Form("CURSO") = "0" then %>
			<TD width="12%"><FONT class="TextoDestaque">&nbsp;CURSO</FONT></TD>
		<% End if %>
		<TD width="30%"><FONT class="TextoDestaque">&nbsp;ÓRGÃO</FONT></TD>
		<TD width="8%"><FONT class="TextoDestaque">&nbsp;CHAVE</TD>
		<TD width="45%"><FONT class="TextoDestaque">&nbsp;NOME</B></TD>
		<TD width="10%"><FONT class="TextoDestaque">&nbsp;CONCLUÍDO</B></TD>		
	</TR>
	<% do while not RS_Proj.eof  %>
	   <% Totaluno = Totaluno + 1 %>
	<TR>
	   <%
	      aprov = ""
	      
	      chave = RS_Proj("USMA_CD_USUARIO")
	      sQuery = "SELECT DISTINCT APROV"
	      sQuery = sQuery & " FROM TabInscritos"
	      sQuery = sQuery & " WHERE COD_DISCIPLINA = '" & RS_Proj("COD_DISCIPLINA") & "'"
	      sQuery = sQuery & " AND   CHAVE = '" & chave & "'"	      
	      sQuery = sQuery & " AND   APROV = 'AP'"	      	      

             set RS_aprov=Conn2.Execute(sQuery)	      
           %>
             <% If not RS_aprov.eof Then  %>
                <% aprov = "OK" %>
                <% Totaluno_aprov = Totaluno_aprov + 1 %>
             <%End if %>
	      
		<% if Request.Form("CURSO") = "0" then %>
			<TD><FONT class="TextoTabela">&nbsp<%=RS_Proj("COD_DISCIPLINA")%></TD>	
		<% End if %>
		<TD><FONT class="TextoTabela">&nbsp<%=RS_Proj("ORME_SG_ORG_MENOR")%></TD>	
		<TD><FONT class="TextoTabela">&nbsp<%=RS_Proj("USMA_CD_USUARIO")%></TD>
		<TD><FONT class="TextoTabela">&nbsp<%=RS_Proj("USMA_TX_NOME_USUARIO")%></TD>
		<TD><FONT class="TextoTabela">&nbsp<%= aprov %></TD>		
	</TR>
	        <%  RS_Proj.moveNext %>
        <% Loop %>
  </TABLE>
  </TD>
  <TD>&nbsp;</TD>
  </TR>
   <TR>
  <TD>
  <table>
  <tr>
  <td><FONT class="Titulo">Total de Indicados Prioritários = <%= Totaluno %></td></tr>
  <tr><td><FONT class="Titulo">Total de Prioritários Aprovados = <%= Totaluno_aprov %></td></tr>
  <tr><td><FONT class="Titulo">Total de Prioritários Pendentes = <%= Totaluno - Totaluno_aprov %></td></tr>
  </table>      
  </TD>
  <TD>&nbsp;</TD>
  </TR>

<% Else %>
  <TR>
  <TD>Não existem usuários mapeados para este curso!</TD>
  <TD>&nbsp;</TD>
  </TR>

<%End If%>

</TABLE>
<%
	RS_Proj.Close
	Set RS_Proj = nothing
	Conn.Close
	Set Conn = nothing
	Conn2.Close
	Set Conn2 = nothing
%>
</body>
<SCRIPT>
	//parent.fra_lista.fun_hide_aguarde();
</SCRIPT>
</html>
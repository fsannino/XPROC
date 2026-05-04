<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_UsuaGestor = ""
str_UsuaGestor = str_UsuaGestor & " SELECT  TOP 100 PERCENT "
str_UsuaGestor = str_UsuaGestor & " USMA_CD_USUARIO "
str_UsuaGestor = str_UsuaGestor & " , USMA_TX_NOME_USUARIO "
str_UsuaGestor = str_UsuaGestor & " FROM dbo.USUARIO_MAPEAMENTO "
str_UsuaGestor = str_UsuaGestor & " Where USMA_TX_MATRICULA <> 0"
str_UsuaGestor = str_UsuaGestor & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_UsuaGestor = db_Cogest.Execute(str_UsuaGestor)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<select name="selUsuarioGestor" size="1" class="listResponsavel">
  <option value="">== Selecione um Gestor ==</option>
  <% 
	  'contRegistro = 0
	  rds_UsuaGestor.movefirst
	  Do While not rds_UsuaGestor.Eof 'and contRegistro < 10
		if str_UsuarioGestor = trim(rds_UsuaGestor("USMA_CD_USUARIO")) then%>
			<option value=<%=rds_UsuaGestor("USMA_CD_USUARIO")%> selected><%=rds_UsuaGestor("USMA_TX_NOME_USUARIO")%></option>
		<%else%>
			<option value=<%=rds_UsuaGestor("USMA_CD_USUARIO")%>><%=rds_UsuaGestor("USMA_TX_NOME_USUARIO")%></option>
		<%end if 
		'contRegistro = contRegistro + 1
		rds_UsuaGestor.movenext		  	
	  Loop
	rds_UsuaGestor.close
	set rds_UsuaGestor = Nothing
	%>
</select>

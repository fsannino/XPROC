<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_UsuaResp = ""
str_UsuaResp = str_UsuaResp & " SELECT  TOP 100 PERCENT "
str_UsuaResp = str_UsuaResp & " USMA_CD_USUARIO "
str_UsuaResp = str_UsuaResp & " , USMA_TX_NOME_USUARIO "
str_UsuaResp = str_UsuaResp & " FROM dbo.USUARIO_MAPEAMENTO "
str_UsuaResp = str_UsuaResp & " Where USMA_TX_MATRICULA <> 0"
str_UsuaResp = str_UsuaResp & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_UsuaResp = db_Cogest.Execute(str_UsuaResp)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selUsuarioResponsavel" class="listResponsavel">
  <option value="">== Selecione um Responsável ==</option>
  	  <% 
	  'contRegistro = 0
	  rds_UsuaResp.movefirst
	  Do While not rds_UsuaResp.Eof 'and contRegistro < 10
	  	if str_UsuarioResponsavel = trim(rds_UsuaResp("USMA_CD_USUARIO")) then%>
			<option value=<%=rds_UsuaResp("USMA_CD_USUARIO")%> selected><%=rds_UsuaResp("USMA_TX_NOME_USUARIO")%></option>
		<%else%>
			<option value=<%=rds_UsuaResp("USMA_CD_USUARIO")%>><%=rds_UsuaResp("USMA_TX_NOME_USUARIO")%></option>
		<%end if		
		'contRegistro = contRegistro + 1
		rds_UsuaResp.movenext		  	
		Loop
	rds_UsuaResp.close
	set rds_UsuaResp = Nothing
	%>
</select>

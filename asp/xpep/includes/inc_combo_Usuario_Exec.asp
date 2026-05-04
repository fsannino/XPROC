<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_UsuaExec = ""
str_UsuaExec = str_UsuaExec & " SELECT  TOP 100 PERCENT "
str_UsuaExec = str_UsuaExec & " USMA_CD_USUARIO "
str_UsuaExec = str_UsuaExec & " , USMA_TX_NOME_USUARIO "
str_UsuaExec = str_UsuaExec & " FROM dbo.USUARIO_MAPEAMENTO "
str_UsuaExec = str_UsuaExec & " Where USMA_TX_MATRICULA <> 0"
str_UsuaExec = str_UsuaExec & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_UsuaExec = db_Cogest.Execute(str_UsuaExec)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selUsuarioExec">
  <option value="">== Selecione um executor ==</option>
  <% 
			  'contRegistro = 0
			  rds_UsuaExec.movefirst
			  Do While not rds_UsuaExec.Eof 'and contRegistro < 10%>
  <option value=<%=rds_UsuaExec("USMA_CD_USUARIO")%>><%=rds_UsuaExec("USMA_TX_NOME_USUARIO")%></option>
  <% 
				'contRegistro = contRegistro + 1
				rds_UsuaExec.movenext		  	
		      Loop
		rds_UsuaExec.close
		set rds_UsuaExec = Nothing
		%>
</select>

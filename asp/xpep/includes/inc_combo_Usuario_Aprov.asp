<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_UsuaAprov = ""
str_UsuaAprov = str_UsuaAprov & " SELECT  TOP 100 PERCENT "
str_UsuaAprov = str_UsuaAprov & " USMA_CD_USUARIO "
str_UsuaAprov = str_UsuaAprov & " , USMA_TX_NOME_USUARIO "
str_UsuaAprov = str_UsuaAprov & " FROM dbo.USUARIO_MAPEAMENTO "
str_UsuaAprov = str_UsuaAprov & " Where USMA_TX_MATRICULA <> 0"
str_UsuaAprov = str_UsuaAprov & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_UsuaAprov = db_Cogest.Execute(str_UsuaAprov)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selUsuarioAprovador">
  <option value="">== Selecione um Aprovador ==</option>
  <% 
			  'contRegistro = 0
			  rds_UsuaAprov.movefirst
			  Do While not rds_UsuaAprov.Eof 'and contRegistro < 10%>
  <option value=<%=rds_UsuaAprov("USMA_CD_USUARIO")%>><%=rds_UsuaAprov("USMA_TX_NOME_USUARIO")%></option>
  <% 
				'contRegistro = contRegistro + 1
				rds_UsuaAprov.movenext		  	
		      Loop
		rds_UsuaAprov.close
		set rds_UsuaAprov = Nothing
		%>
</select>

<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_UsuaContr = ""
str_UsuaContr = str_UsuaContr & " SELECT  TOP 100 PERCENT "
str_UsuaContr = str_UsuaContr & " USMA_CD_USUARIO "
str_UsuaContr = str_UsuaContr & " , USMA_TX_NOME_USUARIO "
str_UsuaContr = str_UsuaContr & " FROM dbo.USUARIO_MAPEAMENTO "
str_UsuaContr = str_UsuaContr & " Where USMA_TX_MATRICULA <> 0"
str_UsuaContr = str_UsuaContr & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_UsuaContr = db_Cogest.Execute(str_UsuaContr)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selUsuarioContr">
  <option value="">== Selecione um Controlador ==</option>
  <% 
			  'contRegistro = 0
			  rds_UsuaContr.movefirst
			  Do While not rds_UsuaContr.Eof 'and contRegistro < 10%>
  <option value=<%=rds_UsuaContr("USMA_CD_USUARIO")%>><%=rds_UsuaContr("USMA_TX_NOME_USUARIO")%></option>
  <% 
				'contRegistro = contRegistro + 1
				rds_UsuaContr.movenext		  	
		      Loop
		rds_UsuaContr.close
		set rds_UsuaContr = Nothing
		%>
</select>

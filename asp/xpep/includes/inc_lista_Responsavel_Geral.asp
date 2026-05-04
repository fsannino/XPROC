<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<span class="campo">
<select name="<%=intIndice%>" size="1" class="listResponsavel" id="<%=intIndice%>">
  <option value="0">== Selecione um Respons&aacute;vel ==</option>
  <% 
			  'contRegistro = 0
			  rds_RespLegado.movefirst
			  Do While not rds_RespLegado.Eof 'and contRegistro < 10
			     if rds_RespLegado("USMA_CD_USUARIO") = int_Cd_Usuario then %>
  <option value=<%=rds_RespLegado("USMA_CD_USUARIO")%> selected><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
				 <% else %>
  <option value=<%=rds_RespLegado("USMA_CD_USUARIO")%>><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
  <% 
  				 end if
				'contRegistro = contRegistro + 1
				rds_RespLegado.movenext		  	
		      Loop
		'rds_RespLegado.close
		'set rds_RespLegado = Nothing
		%>
</select>
</span>
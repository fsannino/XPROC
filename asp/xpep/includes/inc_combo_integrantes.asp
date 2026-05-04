<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USUA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USUA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO "
str_RespLegado = str_RespLegado & " ORDER BY USUA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">


<select name="lstRespTecSinGeral" size="1" class="listResponsavel" id="lstRespTecSineGeral">
  <% Do While not rds_RespLegado.Eof %>
  <option value=<%=rds_RespLegado("USUA_CD_USUARIO")%>><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
  <% rds_RespLegado.movenext
		         Loop
		rds_RespLegado.close
		set rds_RespLegado = Nothing
		%>
</select>

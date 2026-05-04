<%
str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " SIST_NR_SEQUENCIAL_SISTEMA_LEGADO "
str_RespLegado = str_RespLegado & " , SIST_TX_CD_SISTEMA "
str_RespLegado = str_RespLegado & " , SIST_TX_DESC_SISTEMA_LEGADO "
str_RespLegado = str_RespLegado & " FROM XPEP_SISTEMA_LEGADO "
str_RespLegado = str_RespLegado & " Where SIST_NR_SEQUENCIAL_SISTEMA_LEGADO <> 0"
str_RespLegado = str_RespLegado & " ORDER BY SIST_TX_DESC_SISTEMA_LEGADO "

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<span class="campo">
<%if str_Acao <> "C" then
	set rds_RespLegado = db_Cogest.Execute(str_RespLegado)%>
	
	<select name="<%=intIndice%>" size="1" class="listResponsavel" id="<%=intIndice%>">
		<option value="0">== Selecione um Sistema ==</option>
		<%'contRegistro = 0
		  rds_RespLegado.movefirst
		  Do While not rds_RespLegado.Eof ' and contRegistro < 10
			 if rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO") = cint(int_Cd_Sistema_Legado) then%>
				<option value=<%=rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")%> selected><%=rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO") & " - " & rds_RespLegado("SIST_TX_DESC_SISTEMA_LEGADO")%></option>
			 <%else%>
				<option value=<%=rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")%>><%=rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO") & " - " & rds_RespLegado("SIST_TX_DESC_SISTEMA_LEGADO")%></option>
			<%end if
			'contRegistro = contRegistro + 1
			rds_RespLegado.movenext		  	
		  Loop
		rds_RespLegado.close
		set rds_RespLegado = Nothing
		%>
	</select>
<%else
	str_RespLegado = ""
	str_RespLegado = str_RespLegado & " SELECT SIST_NR_SEQUENCIAL_SISTEMA_LEGADO "
	str_RespLegado = str_RespLegado & " , SIST_TX_CD_SISTEMA "
	str_RespLegado = str_RespLegado & " , SIST_TX_DESC_SISTEMA_LEGADO "
	str_RespLegado = str_RespLegado & " FROM XPEP_SISTEMA_LEGADO "
	str_RespLegado = str_RespLegado & " Where SIST_NR_SEQUENCIAL_SISTEMA_LEGADO = " & cint(int_Cd_Sistema_Legado)	
	set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
	
	'*** Imprime o valor do BD para Consulta - Sistema Legado
	Response.write rds_RespLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO") & " - " & rds_RespLegado("SIST_TX_DESC_SISTEMA_LEGADO")
	
end if%>
</span>
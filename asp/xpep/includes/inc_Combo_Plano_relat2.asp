<%
dim str_Cd_Plano, str_Cd_Fase, k

str_Cd_Plano = request("selPlano")

'response.Write("str_Cd_Plano - " & str_Cd_Plano)
'response.End()

if str_Cd_Plano <> 	"" then
   vet_Sigla_Plano = split(str_Cd_Plano,"|")
   str_Sigla_Plano = vet_Sigla_Plano(2)
end if   
'response.Write(str_Sigla_Plano)
'response.End()

'str_Cd_Fase = request("selFases")
'response.Write(str_Cd_Plano)
'response.end()	
k = 1

str_Sql_Plano = ""
str_Sql_Plano = str_Sql_Plano & " SELECT PLAN_TX_SIGLA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_NR_SEQUENCIA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_DESCRICAO_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_IDENTACAO"
str_Sql_Plano = str_Sql_Plano & " FROM XPEP_PLANO_ENT_PRODUCAO"
str_Sql_Plano = str_Sql_Plano & " WHERE PLAN_NR_SEQUENCIA_PLANO > 0 "

if str_Sigla_Plano <>  "" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_TX_SIGLA_PLANO = '" & str_Sigla_Plano & "' "
end if

if str_CD_Onda <> "" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = " & str_CD_Onda
else
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = 0"
end if

if str_Cd_Fases <> "" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = " & str_Cd_Fases
else	
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = 0"
end if

str_Sql_Plano = str_Sql_Plano & " ORDER BY PLAN_TX_DESCRICAO_PLANO"

'Response.WRITE (str_Sql_Plano)
'RESPONSE.END

set rds_Plano=db_Cogest.execute(str_Sql_Plano)
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">  
<select name="selPlano" size="1" class="cmdPlano" onChange="javascript:chamapagina()">
	<% if rds_Plano.EOF then %>
	<option value="">Plano PAI</option>
	<% str_Cd_Plano = ""
	end if %>
    <%DO UNTIL rds_Plano.EOF=TRUE
      
	  ' if (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PAC") and (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PCE") and (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PCM") then
		' if TRIM(str_Cd_Plano)=trim(rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")) then
		  	%>
		  <%'else%>
				<option selected value=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")%>><%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%></option>
			<%
		 'end if
		'end if
		str_Cd_Plano = rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")
		rds_Plano.MOVENEXT
		LOOP
		rds_Plano.close
		set rds_Plano = Nothing
		%>
</select>

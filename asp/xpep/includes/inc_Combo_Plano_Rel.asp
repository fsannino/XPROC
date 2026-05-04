<%
dim str_Cd_Plano, str_Cd_Fase, k

if request("selPlano") <> "" then
	str_Cd_Plano = request("selPlano")
end if

if str_Cd_Plano <> 	"" then
   vet_Sigla_Plano = split(str_Cd_Plano,"|")
   str_Sigla_Plano = vet_Sigla_Plano(2)
end if   

k = 1

str_Sql_Plano = ""
str_Sql_Plano = str_Sql_Plano & " SELECT PLAN_TX_SIGLA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_NR_SEQUENCIA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_DESCRICAO_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_IDENTACAO"
str_Sql_Plano = str_Sql_Plano & " FROM XPEP_PLANO_ENT_PRODUCAO"
str_Sql_Plano = str_Sql_Plano & " WHERE PLAN_NR_SEQUENCIA_PLANO > 0 "

if str_CD_Onda <> "" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = " & str_CD_Onda
else
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = 0"
end if

if str_CD_Onda = 5 or str_CD_Onda = 7 then
	if str_Cd_Fases <> "" then
		str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = " & str_Cd_Fases
	else	
		str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = 0"
	end if
end if

str_Sql_Plano = str_Sql_Plano & " ORDER BY PLAN_TX_DESCRICAO_PLANO"

set rds_Plano=db_Cogest.execute(str_Sql_Plano)
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">  
<select name="selPlano" size="1" class="cmdPlano" onChange="javascript:chamapagina()">
<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
		<option value="">== Selecione o Plano ==</option>
<% else %>    
	<%'if j <> 1 then%>
		<option value="">== Todas os Planos ==</option>		
<% end if %>
    <%DO UNTIL rds_Plano.EOF=TRUE
      
	   if (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PAC") and (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PCE") and (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PCM") then
		 if TRIM(str_Cd_Plano)=trim(rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")) then
		  	%>
				<option selected value=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")%>><%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%></option>
		  <%else%>
				<option value=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & rds_Plano("PLAN_TX_SIGLA_PLANO")%>><%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%></option>
			<%
		 end if
		end if
		rds_Plano.MOVENEXT
		LOOP
		rds_Plano.close
		set rds_Plano = Nothing
		%>
</select>

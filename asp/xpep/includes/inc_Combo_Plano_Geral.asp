<%

'response.Write("   Sess - " & session("CD_Plano2"))
'response.Write("   var - " & str_Cd_Plano2)
'response.Write("   Req - " & request("selPlano2"))

if request("selPlano2") <> "" then
	str_Cd_Plano2 = request("selPlano2")
	session("CD_Plano2") = str_Cd_Plano2
else
	str_Cd_Plano2 = session("CD_Plano2")
end if
'response.End()
k = 1

str_Sql_Plano = ""
str_Sql_Plano = str_Sql_Plano & " SELECT PLAN_TX_SIGLA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_NR_SEQUENCIA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_DESCRICAO_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_IDENTACAO"
str_Sql_Plano = str_Sql_Plano & " FROM XPEP_PLANO_ENT_PRODUCAO"
str_Sql_Plano = str_Sql_Plano & " WHERE PLAN_NR_SEQUENCIA_PLANO > 0 "
if str_Sigla_Plano_Selecionado = "PCM" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_TX_SIGLA_PLANO <> '" & str_Sigla_Plano_Selecionado & "'"
end if	
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

'Response.WRITE (str_Sql_Plano)
'RESPONSE.END

set rds_Plano=db_Cogest.execute(str_Sql_Plano)
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">  
<select name="<%=strNomeObj%>" size="1" class="cmdPlano"  onChange="javascript:chamapagina()">
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione o Plano ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="">== Todas os Planos ==</option>		
	<% end if %>   
    <%DO UNTIL rds_Plano.EOF=TRUE
      
	   if (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PAC") and (trim(rds_Plano("PLAN_TX_SIGLA_PLANO")) <> "PCE")then
		 if TRIM(str_Cd_Plano2)=trim(rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO")) then
		  	%>
				<option selected value=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO")%>><%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%></option>
		  <%else%>
				<option value=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO")%>><%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%></option>
			<%
		 end if
		end if
		rds_Plano.MOVENEXT
		LOOP
		rds_Plano.close
		set rds_Plano = Nothing
		%>
</select>

<%
dim str_Cd_Task1, x
x = 1

'Response.write "SelPlano " & Request("selPlano")
'if str_Cd_Plano <> "" then
'	if request("selPlano") <> "" then
'		str_Cd_Plano = request("selPlano")
'	end if
'end if
'if Request("selPlano") <> "" then

Dim strValorPlano, str_Identacao
'response.Write("Request---" & Request("selPlano") & " str " & str_Cd_Plano )
'response.End()
if Request("selPlano") <> "" then
	strValorPlano =  split(Request("selPlano"), "|")	
	str_Cd_Plano = strValorPlano(0)
	str_Identacao = strValorPlano(1)
else
	if str_Cd_Plano <> "" then
		strValorPlano =  split(str_Cd_Plano, "|")	
		str_Cd_Plano = strValorPlano(0)
		str_Identacao = strValorPlano(1)
	else
		str_Cd_Plano = ""
		str_Identacao = ""
	end if
end if

'*** PEGA NR DO PLANO DO PROJECT
str_TpPlano = ""
str_TpPlano = str_TpPlano & "Select PLAN_TX_SIGLA_PLANO, PLAN_NR_CD_PROJETO_PROJECT "
str_TpPlano = str_TpPlano & " From XPEP_PLANO_ENT_PRODUCAO "
str_TpPlano = str_TpPlano & " WHERE "
str_TpPlano = str_TpPlano & " PLAN_NR_SEQUENCIA_PLANO = " & str_Cd_Plano
'response.Write str_TpPlano & "<br>"
'RESPONSE.End()

if str_Cd_Plano <> "" then
	set rdsTpPlano = db_Cogest.Execute(str_TpPlano)
	if not rdsTpPlano.Eof then
	   int_Cd_Projeto_Project2 = rdsTpPlano("PLAN_NR_CD_PROJETO_PROJECT")   
	else
	   int_Cd_Projeto_Project2 = ""
	end if
	rdsTpPlano.close
	set rdsTpPlano = Nothing
	'response.Write "int_Cd_Projeto_Project - " & int_Cd_Projeto_Project & "<br>"
	'RESPONSE.End()
end if

str_Cd_Task1 = request("selTask1")

str_Sql_Task = ""
str_Sql_Task = str_Sql_Task & " SELECT   "
str_Sql_Task = str_Sql_Task & " TASK_UID"
str_Sql_Task = str_Sql_Task & " , TASK_NAME"
str_Sql_Task = str_Sql_Task & " , RESERVED_DATA"
str_Sql_Task = str_Sql_Task & " , TASK_START_DATE"
str_Sql_Task = str_Sql_Task & " , TASK_FINISH_DATE"
str_Sql_Task = str_Sql_Task & " FROM MSP_TASKS"
str_Sql_Task = str_Sql_Task & " WHERE (LEN(TASK_OUTLINE_NUM) = 11 or LEN(TASK_OUTLINE_NUM) = 12)"
if str_Identacao <> "" then
	str_Sql_Task = str_Sql_Task & " AND TASK_OUTLINE_NUM LIKE '" & TRIM(str_Identacao) & "%'"
	str_Sql_Task = str_Sql_Task & " AND PROJ_ID = " & int_Cd_Projeto_Project2
else
	str_Sql_Task = str_Sql_Task & " AND TASK_OUTLINE_NUM = '99999'"
	str_Sql_Task = str_Sql_Task & " AND PROJ_ID = 0"
end if
str_Sql_Task = str_Sql_Task & " ORDER BY TASK_NAME"

'response.Write(str_Sql_Task)
'response.End()

set rds_Task=db_Cronograma.execute(str_Sql_Task)

'end if
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
  
<select name="selTask1" size="1" class="cmdTask"  onChange="javascript:chamapagina()">
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione a Atividade ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="">== Todas as Atividades ==</option>		
	<% end if %>   	
	<%'if Request("selPlano") <> "" then		
		Do until rds_Task.eof=true
			 If Trim(Request("selTask1")) = Trim(rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")) then
		  %>
		<option selected value=<%=rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")%>><%=rds_Task("TASK_NAME")%></option>
			<%else%>
		<option value=<%=rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")%>><%=rds_Task("TASK_NAME")%></option>
		<%
			end if
			rds_Task.movenext
		  Loop
		  rds_Task.close
		  set rds_Task = Nothing
			
	'end if	%>
  </select>

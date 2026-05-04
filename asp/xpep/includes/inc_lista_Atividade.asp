<%
str_Cd_Task1 = request("selTask1")
'response.Write(str_Cd_Task1)
str_Sql_Task = ""
str_Sql_Task = str_Sql_Task & " SELECT   "
str_Sql_Task = str_Sql_Task & " TASK_UID"
str_Sql_Task = str_Sql_Task & " , TASK_NAME"
str_Sql_Task = str_Sql_Task & " FROM MSP_TASKS"
str_Sql_Task = str_Sql_Task & " WHERE LEN(TASK_OUTLINE_NUM) = 9"
if str_Cd_Project <> "" then
   str_Sql_Task = str_Sql_Task & " AND PROJ_ID = " & str_Cd_Project
end if   
if str_Cd_Fases <> "" then
'   str_Sql_Task = str_Sql_Task & " AND PROJ_ID = " & str_Cd_Project
end if
str_Sql_Task = str_Sql_Task & " ORDER BY TASK_NAME"
'response.Write(str_Sql_Task)
'response.End()
set rds_Task=db_Cronograma.execute(str_Sql_Task)
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
  
<select name="selTask1" size="5" class="cmdTask"  onChange="javascript:chamapagina()">
  <option value="0">== Selecione o Projeto ==</option>
    <option value="0">== Todas os Projetos ==</option>		
    <%Do until rds_Task.eof=true
         If Trim(str_Cd_Task1) = Trim(rds_Task("TASK_UID")) then
      %>
    <option selected value=<%=rds_Task("TASK_UID")%>><%=rds_Task("TASK_NAME")%></option>
        <%else%>
    <option value=<%=rds_Task("TASK_UID")%>><%=rds_Task("TASK_NAME")%></option>
    <%
		end if
		rds_Task.movenext
	  Loop
	  rds_Task.close
	  set rds_Task = Nothing
		%>
  </select>

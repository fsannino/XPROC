<%
Dimstr_Cd_Project, w
str_Cd_Project = request("selProject")
w = 1

str_Sql_Projetos = ""
str_Sql_Projetos = str_Sql_Projetos & " SELECT   "
str_Sql_Projetos = str_Sql_Projetos & " PROJ_ID"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_NAME"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_AUTHOR"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_COMPANY"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_INFO_CAL_NAME"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_SUBJECT"
str_Sql_Projetos = str_Sql_Projetos & ", PROJ_PROP_TITLE"
str_Sql_Projetos = str_Sql_Projetos & " FROM MSP_PROJECTS"
'response.Write(str_Sql_Projetos)
'response.End()
set rds_Projeto=db_Cronograma.execute(str_Sql_Projetos)
%>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
  
<select name="selProject" size="1" class="cmdPlano" onChange="javascript:chamapagina()">
  <option value="">== Selecione o Projeto ==</option>
  
  	<%if w <> 1 then%>
    	<option value="0">== Todos os Projetos ==</option>
	<%end if%>	
		
    <%DO UNTIL rds_Projeto.EOF=TRUE
      IF TRIM(str_Cd_Project)=trim(rds_Projeto("PROJ_ID")) then
      %>
    <option selected value=<%=rds_Projeto("PROJ_ID")%>><%=rds_Projeto("PROJ_NAME")%></option>
    <%else%>
    <option value=<%=rds_Projeto("PROJ_ID")%>><%=rds_Projeto("PROJ_NAME")%></option>
    <%
		END IF
		rds_Projeto.MOVENEXT
		LOOP
		rds_Projeto.close
		set rds_Projeto = Nothing
		%>
  </select>

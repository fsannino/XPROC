<%
Dim str_Cd_Onda, j
j = 1

'if session("CD_Onda") = "" then	
if Request("selOnda") <> "" then	
	str_Cd_Onda = Request("selOnda")
	Session("CD_Onda") = str_Cd_Onda
else
	str_CD_Onda = Trim(Session("CD_Onda"))
end if

str_Sql_Onda = ""
str_Sql_Onda = str_Sql_Onda & " SELECT ONDA_TX_DESC_ONDA "
str_Sql_Onda = str_Sql_Onda & " , ONDA_CD_ONDA, ONDA_TX_ABREV_ONDA "
str_Sql_Onda = str_Sql_Onda & " FROM ONDA "
str_Sql_Onda = str_Sql_Onda & " WHERE ONDA_CD_ONDA <> 4 "
str_Sql_Onda = str_Sql_Onda & " AND ONDA_CD_ONDA IN (" & Session("AcessoUsuario") & ")"
str_Sql_Onda = str_Sql_Onda & " ORDER BY ONDA_TX_DESC_ONDA"
'
'response.Write(str_Sql_Onda)
'response.End()

'set rds_onda = db_CogestP.execute(str_Sql_Onda)
set rds_onda = db_Cogest.execute(str_Sql_Onda)
%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selOnda" size="1" class="cmdOnda" onChange="javascript:chamapagina()">
<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
		<option value="">== Selecione a Onda ==</option>
<% else %>    
	<%'if j <> 1 then%>
		<option value="">== Todas as Ondas ==</option>		
<% end if %>
	
	<%
	do until rds_onda.EOF = True
		
		if trim(rds_onda("ONDA_CD_ONDA")) <> "1" and trim(rds_onda("ONDA_CD_ONDA")) <> "2" then 		
			if TRIM(str_Cd_Onda)=trim(rds_onda("ONDA_CD_ONDA")) then%>
			<option selected value=<%=rds_onda("ONDA_CD_ONDA")%>><%=rds_onda("ONDA_TX_DESC_ONDA")%></option>
			<%else%>
			<option value=<%=rds_onda("ONDA_CD_ONDA")%>><%=rds_onda("ONDA_TX_DESC_ONDA")%></option>
			<%
			end if
		end if
	rds_onda.MOVENEXT
	loop
	rds_onda.close
	set rds_onda = Nothing
	%>
  </select>

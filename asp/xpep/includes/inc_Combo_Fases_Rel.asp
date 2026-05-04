<%
Dim str_Cd_Fases, i

str_Cd_Fases = Request("selFases")

i = 1
%>

<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selFases" size="1" class="cmdOnda" onChange="javascript:chamapagina()">
  <%if str_Cd_Fases = "0" then%>  
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione uma Fase ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="">== Todas as Fase ==</option>		
	<% end if %>
	<option value="1">Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%elseif str_Cd_Fases = "1" then%>
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione uma Fase ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="0">== Todas as Fase ==</option>		
	<% end if %>
	<option value="1" selected>Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%elseif str_Cd_Fases = "2" then%>
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione uma Fase ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="0">== Todas as Fase ==</option>		
	<% end if %>
	<option value="1">Fase 1</option>	
	<option value="2" selected>Fase 2</option>	
  <%else%>
	<% if Request("pAcao") = "X" then 'edição INCLUIR/ALTERAR %>
			<option value="">== Selecione uma Fase ==</option>
	<% else %>    
		<%'if j <> 1 then%>
			<option value="0">== Todas as Fase ==</option>		
	<% end if %>
	<option value="1">Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%end if%> 
</select>

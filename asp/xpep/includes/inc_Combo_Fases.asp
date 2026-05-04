<%
Dim str_Cd_Fases, i

'if session("CD_Fase") = "" then	
if Request("selFases") <> "" then	
	str_Cd_Fases = Request("selFases")
	session("CD_Fase") = str_Cd_Fases
else
	str_Cd_Fases = session("CD_Fase")
end if

i = 1
%>

<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<select name="selFases" size="1" class="cmdOnda" onChange="javascript:chamapagina()">
  <%if str_Cd_Fases = "0" then%> 	
	<option value="">== Selecione uma Fase ==</option>	
	<option value="1">Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%elseif str_Cd_Fases = "1" then%>	
	<option value="">== Selecione uma Fase ==</option>
	<option value="1" selected>Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%elseif str_Cd_Fases = "2" then%>	
	<option value="">== Selecione uma Fase ==</option>	
	<option value="1">Fase 1</option>	
	<option value="2" selected>Fase 2</option>	
  <%else%>	
	<option value="">== Selecione uma Fase ==</option>	
	<option value="1">Fase 1</option>	
	<option value="2">Fase 2</option>	
  <%end if%> 
</select>

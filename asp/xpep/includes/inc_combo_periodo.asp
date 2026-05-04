 <select name="selUnidMedida" size="1" class="cmd150">          
  <%if str_txtUnidTempo = "Hora" then%>
	<option value="Hora" selected>Hora</option>
  <%else%>
	<option value="Hora">Hora</option>
  <%end if%>
  
  <%if str_txtUnidTempo = "Dia Útil" then%>
	<option value="Dia Útil" selected>Dia Útil</option>
  <%else%>
	<option value="Dia Útil">Dia Útil</option>
  <%end if%>
  
  <%if str_txtUnidTempo = "Dia Corrido" then%>
	<option value="Dia Corrido" selected>Dia Corrido</option>
  <%else%>
	<option value="Dia Corrido">Dia Corrido</option>
  <%end if%>
  
   <%if str_txtUnidTempo = "Mês" then%>
	<option value="Mês" selected>Mês</option>
  <%else%>
	<option value="Mês">Mês</option>
  <%end if%> 
</select>        

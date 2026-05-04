<%
strMega = request("selMega")
strAbr = request("selAbrangencia")
strStatus = request("selStatus")

str_Arq_Imp = "imprime_consulta_curso.asp"
str_Arq_Imp = str_Arq_Imp & "?selMega=" & strMega  & "&selAbrangencia=" & strAbr & "&selStatus=" & strStatus 

'Response.write str_Arq_Imp
%>
<HTML>
<HEAD>
<title>:: Demostrativo de Cursos</title>
</HEAD>
    
<FRAMESET ROWS="100%,2" FRAMEBORDER="1"  border="2" framespacing="0" cols="*"> 
  <FRAME src="msg_imprimindo.asp" name="frame1">
  <FRAME src="<%=str_Arq_Imp%>" name="frame2" noresize>
</FRAMESET>
<noframes></noframes>	
</HTML>





<%'--- Página com o TAMANHO DA NOVA JANELA ---'%>
<%
str_Arq_Imp = Request("par_PaginaPrint")
' noresize
%>
<HTML>
<HEAD>
<title>Imprimindo</title>
</HEAD>   
    <FRAMESET ROWS="50%,0"  FRAMEBORDER="1">
		<FRAME src="msg_aguardando.asp" name="frame1">
		<FRAME src="<%=str_Arq_Imp%>" name="frame2">
	</FRAMESET><noframes></noframes>	
</HTML>





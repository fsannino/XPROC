<%'--- Página com o TAMANHO DA NOVA JANELA ---'%>
<HTML>
<HEAD>
<title>Imprimindo</title>
</HEAD>

    <%
      dim strPaginaPrint
      strPaginaPrint = "teste_print.asp" 'Request("par_PaginaPrint")
    %>

    <FRAMESET ROWS="50%,0"  FRAMEBORDER="1">
		<FRAME src="imprimindo.asp" name="frame1" >
		<FRAME src="<%=strPaginaPrint%>" name="frame2">
	</FRAMESET>	
</HTML>





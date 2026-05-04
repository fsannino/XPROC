<%'--- Página com o TAMANHO DA NOVA JANELA ---'%>
<%

int_Onda = request("selOnda")
int_Fase = request("selFases")
int_Plano1 = request("selPlano")
int_Plano2 = request("selPlano2")
int_Atividade = request("selTask1")
int_AtividadeSub = request("selTaskSub")

str_Arq_Imp = Request("par_PaginaPrint")
str_Arq_Imp = str_Arq_Imp & "?selOnda=" & int_Onda 
str_Arq_Imp = str_Arq_Imp & "&selFases=" & int_Fase 
str_Arq_Imp = str_Arq_Imp & "&selPlano=" & int_Plano1 
str_Arq_Imp = str_Arq_Imp & "&selPlano2=" & int_Plano2 
str_Arq_Imp = str_Arq_Imp & "&selTask1=" & int_Atividade
str_Arq_Imp = str_Arq_Imp & "&selTaskSub=" & int_AtividadeSub

'response.write int_Onda 
'response.write int_Fase 
'response.write int_Plano1 
'response.write int_Plano2 
'response.write int_Atividade 
'response.write int_AtividadeSub
'response.write str_Arq_Imp 
'response.end()
%>
<HTML>
<HEAD>
<title>Imprimindo</title>
</HEAD>   

    <FRAMESET ROWS="50%,0"  FRAMEBORDER="1">
		<FRAME src="msg_imprimindo.asp" name="frame1" >
		<FRAME src="<%=str_Arq_Imp%>" name="frame2" noresize>
	</FRAMESET>	
</HTML>





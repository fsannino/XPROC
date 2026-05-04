<%
linha = 0
identidade = ucase(request("txtIdent"))
if identidade = "XISPOSTOS" then
%>
<!--#include file="../conecta.asp" -->
<%
	set rs1 = db.execute("SELECT DISTINCT EQUIPE FROM " & Session("Tabela") & " WHERE EQUIPE IN ('', 'BASIS-PETROBRAS', 'GESTÃO DE EMPREENDIMENTOS', 'MANUTENÇÃO/INSPEÇÃO', 'MES-TREINAMENTO', 'ORÇAMENTO','SERVIÇOS', 'TREINAMENTO') ORDER BY EQUIPE")
	
	set mail = Server.CreateObject("Persits.MailSender")
	
	mail.host = "164.85.62.165"
	mail.from = "suporte@S600146.petrobras.com.br"
	mail.FromName = "Suporte Sinergia"
	
	mail.addaddress "xd47@petrobras.com.br"
		
	mail.subject = "Informações para Análise e Validação"
	
	mail.body = "Solicitamos atentar aos seguintes Valores : " & chr(13) & chr(13)
	
	mail.body = mail.body & "================================================"  & chr(13)
	mail.body = mail.body & "========== PARÂMETROS DE CONFIGURAÇÃO =========="  & chr(13)
	mail.body = mail.body & "================================================"  & chr(13) & chr(13)
	mail.body = mail.body & "DATA BASE : " & Session("data_inicio") & chr(13)
	mail.body = mail.body & "PERÍODO : " & Session("periodo") & " dias" & chr(13)
	mail.body = mail.body & "TIPO : " & Session("Erro") & chr(13)
	mail.body = mail.body & "ORGAO : " & Session("Orgao") & chr(13) & chr(13)
	mail.body = mail.body & "================================================"  & chr(13)
	
	do until rs1.eof=true                      
    
    if trim(rs1("EQUIPE"))="" then                      
    	Equipe="ATENDENTE TI"
    else
    	equipe=RS1("EQUIPE")
    end if
	
    data_01 = cdate(session("data_inicio"))
	data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)

	ssql=""
	ssql="SELECT * FROM " & Session("tabela") & " WHERE SITUACAO='ABERTO' AND EQUIPE ='" & rs1("EQUIPE") & "'"
	ssql=ssql+" AND (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
	ssql=ssql+ Session("compl")
			
	set rs2 = db.execute(ssql)
           
    abertos = rs2.recordcount

	ssql=""
	ssql="SELECT * FROM " & Session("tabela") & " WHERE SITUACAO='EM ANDAMENTO' AND EQUIPE ='" & rs1("EQUIPE") & "'"
	ssql=ssql+" AND (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
	ssql=ssql+ Session("compl")
			
	set rs2 = db.execute(ssql)
           
    andamento = rs2.recordcount

	ssql=""
	ssql="SELECT * FROM " & Session("tabela") & " WHERE SITUACAO='PENDENTE' AND EQUIPE ='" & rs1("EQUIPE") & "'"
	ssql=ssql+" AND (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
	ssql=ssql+ Session("compl")
			
	set rs2 = db.execute(ssql)
           
    pendentes = rs2.recordcount
	
	if abertos > 0 or andamento > 0 or pendentes > 0 then
	
		mail.body = mail.body & "================================================"
		mail.body = mail.body & chr(13) & chr(13)
		mail.body = mail.body & "MÓDULO : "& EQUIPE
		mail.body = mail.body & chr(13) & chr(13)
		mail.body = mail.body & "ABERTOS : " & abertos & chr(13)
		mail.body = mail.body & "EM ANDAMENTO : " & andamento & chr(13)
		mail.body = mail.body & "PENDENTES : " & pendentes & chr(13)
		mail.body = mail.body & chr(13)
	
		linha = linha + 1
	
	end if
	
	rs1.movenext
	loop
	
	mail.body = mail.body & "================================================"
	
	if linha > 0 then	
		mail.send	
	end if
		
	enviou=1

else
	enviou=0
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 1</title>

</head>

<script>
function Enviar()
{
if(document.frm1.txtIdent.value=="")
{
alert('Você deve digitar a sua identificação!');
document.frm1.txtIdent.focus();
return;
}
else
{
document.frm1.submit();
}
}
</script>

<%
if enviou=0 then
%>
<body onLoad="document.frm1.txtIdent.focus()" link="#800000" vlink="#800000" alink="#800000">
<%else%>
<body link="#800000" vlink="#800000" alink="#800000">
<%end if%>

<%
if enviou=0 then
%>
<p><b><font face="Verdana" size="2">Digite a sua identificação do Sistema</font></b></p>

<form method="POST" name="frm1" action="correio.asp">
   
   <p><input type="password" name="txtIdent" size="24" maxlength="16"></p>
   
   <p><input type="button" value="Enviar Correio" name="B1" onClick="Enviar()"></p>
   
</form>
<%
else
if linha >0 then
	mensagem="A sua Mensagem foi enviada com Sucesso!"
else
	mensagem="Não existem dados à serem enviados"
end if
%>
<p><b><font face="Verdana" size="2" color="#000080"><%=Mensagem%></font></b></p>

<p><b><font face="Verdana" size="2" color="#000080"><a href="../target.asp">&lt;&lt;&lt; Voltar</a></font></b></p>
<%end if%>
</body>

</html>
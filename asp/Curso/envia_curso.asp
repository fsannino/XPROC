<%
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

opt=request("option")
mega=request("mega")
curso=ucase(request("curso"))

if curso="" then
	curso=ucase(request("mega"))
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")
	if rs.eof=false then
		mega=rs("MEPR_CD_MEGA_PROCESSO")
	else
		response.redirect "seleciona_curso.asp?option=" & opt &"&resp=1"
	end if
end if

select case opt
		
	'Seleção para funções básicas
	
	case 1
		response.redirect "rel_curso_transacao.asp?mega=" & mega & "&curso=" & curso
	case 2
		response.redirect "rel_curso_funcao.asp?mega=" & mega & "&curso=" & curso
	case 3
		response.redirect "rel_curso_cenario.asp?mega=" & mega & "&curso=" & curso
	case 4
		response.redirect "rel_curso_pre_requisitos.asp?mega=" & mega & "&curso=" & curso
	case 6
		response.redirect "altera_curso.asp?curso=" & curso
	case 7
		response.redirect "rel_curso_pre_requisitos_alternativo.asp?mega=" & mega & "&curso=" & curso
	case 8
		response.redirect "rel_curso_correlato.asp?mega=" & mega & "&curso=" & curso

end select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>

<script>
function verifica_excluir(){
if(document.frm1.txtopt.value=5)
{
if(confirm("deseja realmente excluir?(TODOS os registros relacionados serão excluídos!(Relações com Função de Negócio, Cenários, Transações ,Cursos Pre-Requisitos e Alternativos)"))
{
window.location.href="valida_exclui_curso.asp?curso=<%=curso%>";
}
else
{
window.location.href="seleciona_curso.asp?option="+document.frm1.txtopt.value;
}
}
}
</script>

<body onload="javascript:verifica_excluir()">
<form method="POST" action="" name="frm1">
  <p><input type="hidden" name="txtopt" size="20" value="<%=opt%>"></p>
</form>
</body>
</html>

<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

opt			= request("option")
mega		= request("mega")
curso		= ucase(request("curso"))
strStatus 	= request("rdbStatus")

if curso="" then
	curso=ucase(request("mega"))
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")
	if rs.eof=false then
		mega=rs("MEPR_CD_MEGA_PROCESSO")
	else
		response.redirect "seleciona_curso_rel.asp?option=" & opt &"&resp=1"
	end if
end if

if mega <> 0 then
	complemento = "mega=" & mega
end if

if curso <> "" then
	complemento = complemento + "&curso=" & curso
end if

complemento = complemento + "&status=" & strStatus

select case opt
	case 1
		response.redirect "relat_curso_transacao.asp?" & complemento
	case 2
		response.redirect "relat_curso_funcao.asp?" & complemento
	case 3
		response.redirect "relat_curso_cenario.asp?" & complemento
	case 4
		response.redirect "relat_curso_pre_requisito.asp?" & complemento
	case 5
		response.redirect "relat_curso_correlato.asp?" & complemento
end select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>

<body>
<form method="POST" action="" name="frm1">
  <p><input type="hidden" name="txtopt" size="20" value="<%=opt%>"></p>
</form>
</body>
</html>

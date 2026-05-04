<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Chave = Request("chave")
Session("CdUsuario") = str_Chave

if Session("CdUsuario") = "XK45" or Session("CdUsuario") = "XD47" or Session("CdUsuario") = "XT54" or str_Chave = "XK45" or str_Chave = "XD47" or str_Chave = "XT54" then
   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
else 
   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

Session("Prefixo")="dbo."

ls_SQL = ""
ls_SQL = ls_SQL & "SELECT USUA_TX_NOME_USUARIO, USUA_TX_CATEGORIA, USUA_CD_USUARIO "
ls_SQL = ls_SQL & "FROM GRADE_USUARIO "
ls_SQL = ls_SQL & "WHERE USUA_CD_USUARIO = '" & str_chave & "'" 

Set rdsAcesso= conn_db.Execute(ls_SQL)
if rdsAcesso.EOF then  
   Session("CategoriaUsu") = "indexD.htm"
   Session("CatUsu") = "indexD.js"  
else
   Session("NomeUsuario") = rdsAcesso("USUA_TX_NOME_USUARIO")
   Session("CdUsuario") = rdsAcesso("USUA_CD_USUARIO")
   ls_Categoria = rdsAcesso("USUA_TX_CATEGORIA")
   ls_Controle = "0"
   
   Select Case ls_Categoria
	   Case "A"
		  Session("CategoriaUsu") = "indexA.htm"
		  Session("CatUsu") = "indexA.js"
	   Case "B"
		  Session("CategoriaUsu") = "indexB.htm"
		  Session("CatUsu") = "indexB.js"
	   Case "C"
		  Session("CategoriaUsu") = "indexC.htm"
		  Session("CatUsu") = "indexC.js"
	   Case "D"
		  Session("CategoriaUsu") = "indexD.htm"
		  Session("CatUsu") = "indexD.js"	
	   Case "Q"
		  Session("CategoriaUsu") = "indexQ.htm"
		  Session("CatUsu") = "indexQ.js"
   end Select
end if

UrlNova = "indexA_grade.htm"
response.redirect(UrlNova)

rdsAcesso.close
set rdsAcesso = Nothing
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p><%=ls_SQL%></p>
<p><%=UrlNova%></p>
</body>
</html>
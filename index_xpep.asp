<%@LANGUAGE="VBSCRIPT"%> 
<%
Response.Expires=0
str_Chave = Request("chave")
Session("CdUsuario") = str_Chave
'response.Write(str_Chave)
'response.End()

if Session("CdUsuario") = "XK45" or Session("CdUsuario") = "XD47" or Session("CdUsuario") = "XT54" or str_Chave = "XK45" or str_Chave = "XT54" or str_Chave = "XD47" then
   Session("Conn_String_Cronograma_Gravacao")= "Provider=SQLOLEDB.1;server=10.2.56.131;pwd=sinergia;uid=sinergia;database=pmo"
   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
else
   Session("Conn_String_Cronograma_Gravacao")= "Provider=SQLOLEDB.1;server=10.2.56.131;pwd=sinergia;uid=sinergia;database=pmo"
   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

Session("Prefixo")="dbo."

ls_SQL = ""
ls_SQL = ls_SQL & " SELECT " & Session("PREFIXO") & "XPEP_ACESSO.USUA_CD_USUARIO, "
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "XPEP_USUARIO.USUA_TX_NOME_USUARIO, "
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "XPEP_ACESSO.ONDA_CD_ONDA,"
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "XPEP_USUARIO.USUA_TX_CATEGORIA"
ls_SQL = ls_SQL & " FROM " & Session("PREFIXO") & "XPEP_ACESSO INNER JOIN"
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "XPEP_USUARIO ON " & Session("PREFIXO") & "XPEP_ACESSO.USUA_CD_USUARIO = " & Session("PREFIXO") & "XPEP_USUARIO.USUA_CD_USUARIO"
ls_SQL = ls_SQL & " WHERE " & Session("PREFIXO") & "XPEP_ACESSO.USUA_CD_USUARIO = '" & str_chave & "'" 

'Response.write ls_SQL
'Response.end

Session("AcessoUsuario") = ""
Set rdsAcesso= Conn_db.Execute(ls_SQL)
if rdsAcesso.EOF then
   'Session("CdUsuario") = str_chave
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/sem_acesso.htm"
   'response.redirect(UrlNova)
   Session("CategoriaUsu") = "indexD.htm"
   Session("CatUsu") = "indexD.js"
   Session("NomeUsuario") = "PERFIL DE CONSULTA" 
else
   Session("NomeUsuario") = rdsAcesso("USUA_TX_NOME_USUARIO")
   Session("CdUsuario") = rdsAcesso("USUA_CD_USUARIO")
   ls_Categoria = rdsAcesso("USUA_TX_CATEGORIA")
   ls_Controle = "0"
   do while not rdsAcesso.EOF
      if ls_Controle = "0" then
         Session("AcessoUsuario") = rdsAcesso("ONDA_CD_ONDA")
	     ls_Controle = "1"
	  else	 
         Session("AcessoUsuario") = Session("AcessoUsuario") & "," & rdsAcesso("ONDA_CD_ONDA")
      end if		 
      rdsAcesso.movenext
   loop
   
   'response.write ls_Categoria
   'Response.End
   
   Select Case ls_Categoria
   Case "A"
	  Session("CategoriaUsu") = "indexA.htm"
	  Session("CatUsu") = "indexA.js"	  
   Case "B"
	  Session("CategoriaUsu") = "indexB.htm"
	  Session("CatUsu") = "indexB.js"	  
   Case "D"
	  Session("CategoriaUsu") = "indexD.htm"
	  Session("CatUsu") = "indexD.js"   
   Case "Q"
	  Session("CategoriaUsu") = "indexQ.htm"
	  Session("CatUsu") = "indexQ.js"

   end Select
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
'   UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
'   response.redirect(UrlNova)
end if

   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA_xpep.htm"

rdsAcesso.close
set rdsAcesso = Nothing

'response.write Session("CatUsu")
'response.write Session("ls_Categoria")
'response.End()

UrlNova = "indexA_xpep.htm"
response.redirect(UrlNova)
'response.write Session("CatUsu")
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

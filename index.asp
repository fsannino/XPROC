<%@LANGUAGE="VBSCRIPT"%>
<%
'response.End()
str_Chave = UCase(Request("chave"))
str_Senha = UCase(Request("senha"))
str_Origem_Link = request("pOrigemLink")
response.Write(str_Chave)
response.Write(str_Senha)
response.Write(pOrigemLink)
'response.End()
Session("chave") = str_Chave
Session("senha") = str_Senha
'Session("CdUsuario") = "01CO" or Session("CdUsuario") = "02CO" or Session("CdUsuario") = "03CO" or Session("CdUsuario") = "04CO" or Session("CdUsuario") = "05CO" or Session("CdUsuario") = "06CO" or Session("CdUsuario") = "07CO" or Session("CdUsuario") = "08CO" or
'Session("CdUsuario") = "09CO" or Session("CdUsuario") = "10CO" or Session("CdUsuario") = "11CO" or Session("CdUsuario") = "12CO" or Session("CdUsuario") = "13CO" or Session("CdUsuario") = "14CO" or Session("CdUsuario") = "15CO" or Session("CdUsuario") = "16CO" or
'Session("CdUsuario") = "17CO" or Session("CdUsuario") = "18CO" or Session("CdUsuario") = "19CO" or Session("CdUsuario") = "20CO" or Session("CdUsuario") = "01FI" or Session("CdUsuario") = "02FI" or Session("CdUsuario") = "03FI" or Session("CdUsuario") = "04FI" or
'Session("CdUsuario") = "05FI" or Session("CdUsuario") = "06FI" or Session("CdUsuario") = "07FI" or Session("CdUsuario") = "08FI" or Session("CdUsuario") = "09FI" or Session("CdUsuario") = "10FI" or Session("CdUsuario") = "11FI" or Session("CdUsuario") = "12FI" or
'Session("CdUsuario") = "13FI" or Session("CdUsuario") = "14FI" or Session("CdUsuario") = "15FI" or Session("CdUsuario") = "16FI" or Session("CdUsuario") = "17FI" or Session("CdUsuario") = "18FI" or Session("CdUsuario") = "19FI" or Session("CdUsuario") = "20FI" or


'Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=localhost;pwd=;uid=sa;database=cogest"
Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server="& request.servervariables("remote_host") & ";pwd=cogest;uid=sa;database=cogest"
'response.write(Session("Conn_String_Cogest_Gravacao"))
'response.end
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

if str_Origem_Link <> "NOTES" then
	ls_SQL = ""
	ls_SQL = ls_SQL & " SELECT "
	ls_SQL = ls_SQL & " USUA_CD_USUARIO"
	ls_SQL = ls_SQL & " ,USAA_TX_SENHA"
	ls_SQL = ls_SQL & " FROM dbo.USUARIO"
	ls_SQL = ls_SQL & " WHERE  USUA_CD_USUARIO = '" & str_Chave & "'"
	set rdsUsuario = Conn_db.Execute(ls_SQL)
	if rdsUsuario.EOF then
		'UrlNova = "msg_erro_geral.asp?pCdMsgErro=1"
		'response.redirect(UrlNova)
		str_chave = "GGGG"					
	else
		if rdsUsuario("USAA_TX_SENHA") <> str_Senha then
			UrlNova = "msg_erro_geral.asp?pCdMsgErro=2" 
			response.Redirect(UrlNova)			
		end if
	
	end if
end if

Session("Prefixo")="dbo."

ls_SQL = ""
ls_SQL = ls_SQL & " SELECT ACESSO.USUA_CD_USUARIO, "
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "USUARIO.USUA_TX_NOME_USUARIO, "
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "ACESSO.MEPR_CD_MEGA_PROCESSO,"
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "USUARIO.USUA_TX_CATEGORIA"
ls_SQL = ls_SQL & " FROM " & Session("PREFIXO") & "ACESSO INNER JOIN"
ls_SQL = ls_SQL & " " & Session("PREFIXO") & "USUARIO ON " & Session("PREFIXO") & "ACESSO.USUA_CD_USUARIO = " & Session("PREFIXO") & "USUARIO.USUA_CD_USUARIO"
ls_SQL = ls_SQL & " WHERE " & Session("PREFIXO") & "ACESSO.USUA_CD_USUARIO = '" & str_chave & "'"

Session("AcessoUsuario") = ""
Set rdsAcesso= Conn_db.Execute(ls_SQL)
if rdsAcesso.EOF then
   'Session("CdUsuario") = str_chave
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/sem_acesso.htm"
   'response.redirect(UrlNova)
   Session("CategoriaUsu") = "indexD.htm"
   Session("CatUsu") = "indexD.js"
else
   Session("NomeUsuario") = rdsAcesso("USUA_TX_NOME_USUARIO")
   Session("CdUsuario") = rdsAcesso("USUA_CD_USUARIO")
   ls_Categoria = rdsAcesso("USUA_TX_CATEGORIA")
   ls_Controle = "0"
   do while not rdsAcesso.EOF
      if ls_Controle = "0" then
         Session("AcessoUsuario") = rdsAcesso("MEPR_CD_MEGA_PROCESSO")
	     ls_Controle = "1"
	  else
         Session("AcessoUsuario") = Session("AcessoUsuario") & "," & rdsAcesso("MEPR_CD_MEGA_PROCESSO")
      end if
      rdsAcesso.movenext
   loop
   'response.write ls_Categoria
   Select Case ls_Categoria
   Case "A"
	  Session("CategoriaUsu") = "indexA.htm"
	  Session("CatUsu") = "indexA.js"
   Case "B"
	  Session("CategoriaUsu") = "indexA.htm"
	  Session("CatUsu") = "indexB.js"
   Case "C"
	  Session("CategoriaUsu") = "indexC.htm"
	  Session("CatUsu") = "indexC.js"
   Case "D"
	  Session("CategoriaUsu") = "indexD.htm"
	  Session("CatUsu") = "indexD.js"
   Case "E"
	  Session("CategoriaUsu") = "indexE.htm"
	  Session("CatUsu") = "indexE.js"
   Case "F"
	  Session("CategoriaUsu") = "indexF.htm"
	  Session("CatUsu") = "indexF.js"
	Case "G"
	  Session("CategoriaUsu") = "indexG.htm"
	  Session("CatUsu") = "indexG.js"
	Case "H"
	  Session("CategoriaUsu") = "indexH.htm"
	  Session("CatUsu") = "indexH.js"
   Case "P"
	  Session("CategoriaUsu") = "indexP.htm"
	  Session("CatUsu") = "indexP.js"
   Case "Q"
	  Session("CategoriaUsu") = "indexQ.htm"
	  Session("CatUsu") = "indexQ.js"
   Case "V"
	  Session("CategoriaUsu") = "indexV.htm"
	  Session("CatUsu") = "indexV.js"
   Case "Z"
	  Session("CategoriaUsu") = "indexZ.htm"
	  Session("CatUsu") = "indexZ.js"
   Case "W"
	  Session("CategoriaUsu") = "indexW.htm"
	  Session("CatUsu") = "indexW.js"

   end Select
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
'   UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
'   response.redirect(UrlNova)
end if

   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
   'UrlNova = "http://S6000WS10.corp.petrobras.biz/xproc/indexA.htm"
   'UrlNova = "indexA.htm"
   'response.redirect(UrlNova)

'response.write Session("CatUsu")

rdsAcesso.close
set rdsAcesso = Nothing

%>

<html>
<head>
<title>X-PROC</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>

<script language="JavaScript" type="text/JavaScript">
<!--
function jump_4()
{
    var sscWindow
    sscWindow= window.open('http://localhost/xproc/indexA.asp', 'test4', 'left=0,top=0,resizable=no,scrollbars=yes,fullscreen=yes,toolbar=no,location=no');

    if (window.focus)
    {
        sscWindow.focus()
    }
    return false;
}
//-->


<!--
    opener.opener = opener;
    opener.close();
//-->
</script>
<body bgcolor="#FFFFFF" text="#000000" onload="javascript:jump_4(); window.close()">

<p><%'=ls_SQL%> </p>
<!--<p><%=UrlNova%> </p>-->
</body>
</html>
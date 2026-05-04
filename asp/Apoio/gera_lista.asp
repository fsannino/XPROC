<%@LANGUAGE="VBSCRIPT"%> 
<%
StrConn = "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"

Response.Buffer=false

server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open StrConn

response.write "Acessando..." & strConn & "<p>"

set fs = server.CreateObject("Scripting.FileSystemObject")

set rs = db.execute("SELECT DISTINCT USMA_CD_USUARIO FROM APOIO_LOCAL_MULT WHERE APLO_NR_ATRIBUICAO=1 AND APLO_NR_SITUACAO=1 ORDER BY USMA_CD_USUARIO")

caminho1 = Server.Mappath("../../Publico/Lista/Alis.txt")
set arquivo = fs.CreateTextFile(caminho1)

do until rs.eof=true
	ali = ali & ucase(TRIM(rs("usma_cd_usuario"))) & ","	
	rs.movenext
loop

arquivo.writeline left(ali, len(ali)-1)
Response.Write "<br>===>Alis.txt Gerada "

set rs = db.execute("SELECT DISTINCT USMA_CD_USUARIO FROM CLI ORDER BY USMA_CD_USUARIO")

caminho2 = Server.Mappath("../../Publico/Lista/Clis.txt")
set arquivo = fs.CreateTextFile(caminho2)

do until rs.eof=true
	cli = cli & ucase(TRIM(rs("usma_cd_usuario"))) & ","	
	rs.movenext
loop

arquivo.writeline left(cli, len(cli)-1)
Response.Write "<br>===>Clis.txt Gerada "

ssql=""
ssql="SELECT DISTINCT DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
ssql=ssql+"FROM APOIO_LOCAL_MULT "
ssql=ssql+"INNER JOIN APOIO_LOCAL_ORGAO ON "
ssql=ssql+"DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO = DBO.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
ssql=ssql+"WHERE DBO.APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1 AND DBO.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1  AND DBO.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO=1 "
ssql=ssql+"AND DBO.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '88%' ORDER BY DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO"

set rs = db.execute(ssql)

caminho3 = Server.Mappath("../../Publico/Lista/Alis_Abast.txt")
set arquivo = fs.CreateTextFile(caminho3)

do until rs.eof=true
	abast = abast & ucase(TRIM(rs("usma_cd_usuario"))) & ","	
	rs.movenext
loop

arquivo.writeline left(abast, len(abast)-1)
Response.Write "<br>===>Alis_Abast.txt Gerada "

ssql=""
ssql="SELECT DISTINCT DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
ssql=ssql+"FROM APOIO_LOCAL_MULT "
ssql=ssql+"INNER JOIN APOIO_LOCAL_ORGAO ON "
ssql=ssql+"DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO = DBO.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
ssql=ssql+"WHERE DBO.APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1 AND DBO.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1  AND DBO.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO=1 "
ssql=ssql+"AND DBO.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '87%' ORDER BY DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO"

set rs = db.execute(ssql)

caminho4 = Server.Mappath("../../Publico/Lista/Alis_EeP.txt")
set arquivo = fs.CreateTextFile(caminho4)

do until rs.eof=true
	eep = eep & ucase(TRIM(rs("usma_cd_usuario"))) & ","	
	rs.movenext
loop

arquivo.writeline left(eep, len(eep)-1)
Response.Write "<br>===>Alis_EeP.txt Gerada "

ssql=""
ssql="SELECT DISTINCT DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
ssql=ssql+"FROM APOIO_LOCAL_MULT "
ssql=ssql+"INNER JOIN APOIO_LOCAL_ORGAO ON "
ssql=ssql+"DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO = DBO.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
ssql=ssql+"WHERE DBO.APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1 AND DBO.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1  AND DBO.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO=1 "
ssql=ssql+"AND DBO.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR NOT LIKE '87%' AND DBO.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR NOT LIKE '88%' ORDER BY DBO.APOIO_LOCAL_MULT.USMA_CD_USUARIO"

set rs = db.execute(ssql)

caminho5 = Server.Mappath("../../Publico/Lista/Alis_Outros.txt")
set arquivo = fs.CreateTextFile(caminho5)

do until rs.eof=true
	outros = outros & ucase(TRIM(rs("usma_cd_usuario"))) & ","	
	rs.movenext
loop

arquivo.writeline left(outros, len(outros)-1)
Response.Write "<br>===>Alis_Outros.txt Gerada "

set arquivo = nothing
set fs = nothing

set mail = Server.CreateObject("Persits.MailSender")

mail.host = "164.85.62.165"
mail.from = "Sinergia@S600146.petrobras.com.br"
mail.FromName = "Base de Ali/Cli"

mail.addaddress "pe10@petrobras.com.br"
mail.addaddress "xt52@petrobras.com.br"
mail.addaddress "xu65@petrobras.com.br"
mail.addaddress "bs48@petrobras.com.br"
mail.addaddress "xr97@petrobras.com.br"

mail.subject = "Listas de Distribuição"

mail.AddAttachment caminho1
mail.AddAttachment caminho2
mail.AddAttachment caminho3
mail.AddAttachment caminho4
mail.AddAttachment caminho5

mail.send
Response.Write "<p>===>Correio Enviado "

set fs = server.CreateObject("Scripting.FileSystemObject")

fs.Deletefile caminho1
fs.Deletefile caminho2
fs.Deletefile caminho3
fs.Deletefile caminho4
fs.Deletefile caminho5

set fs = nothing
Response.Write "<p>===>Procedimento Concluído!"
%>

<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
</body>
</html>
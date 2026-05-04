<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html><head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<!--#include file="../Adovbs.inc" -->
<% 
'connectme="DSN=Student;uid=student;pwd=magic"
sqltemp="select * from publishers"

int_Cod_Escopo = request("pCodEscopo")

'set conn_db = Server.CreateObject("ADODB.Connection")
'conn_db.Open Session("Conn_String_Cogest_Gravacao")

sqltemp = ""
sqltemp = sqltemp & " SELECT DISTINCT "
sqltemp = sqltemp & " RELACAO_FINAL.TRAN_CD_TRANSACAO, "
sqltemp = sqltemp & " TRANSACAO.TRAN_TX_DESC_TRANSACAO"
sqltemp = sqltemp & " FROM RELACAO_FINAL INNER JOIN"
sqltemp = sqltemp & " TRANSACAO ON "
sqltemp = sqltemp & " RELACAO_FINAL.TRAN_CD_TRANSACAO = TRANSACAO.TRAN_CD_TRANSACAO"
sqltemp = sqltemp & " WHERE FEES_CD_FECHAMENTO_ESCOPO = " & int_Cod_Escopo

If aduseclient="" THEN
ref="http://www.learnasp.com/adovbs.inc"
response.write "You forgot to include:<br>"
response.write "/adovbs.inc<br>"
response.write "Get the file from <a href='" & ref & "'>" & ref & "<br>"
response.end
END IF
mypage=request("whichpage")
If mypage="" then
mypage=1
end if
mypagesize=request("pagesize")
If mypagesize="" then
mypagesize=10
end if
mySQL=request("SQLquery")
IF mySQL="" THEN
mySQL=SQLtemp
END IF
set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.cursorlocation=aduseclient
rstemp.cachesize=5
tempSQL=lcase(mySQL)
badquery=false
IF instr(tempSQL,"delete")>0 THEN
badquery=true
END IF
IF instr(tempSQL,"insert")>0 THEN
badquery=true
END IF
IF instr(tempSQL,"update")>0 THEN
badquery=true
END IF
If badquery=true THEN
response.write "Not a SELECT Statement<br>"
response.end
END IF
'response.Write(mySQL)
'response.Write("  -   ")
'response.Write(Conn_String_Cogest_Gravacao)
rstemp.open mySQL,Session("Conn_String_Cogest_Gravacao")

if not rstemp.EOF then

rstemp.movefirst
rstemp.pagesize=mypagesize
maxpages=cint(rstemp.pagecount)
maxrecs=cint(rstemp.pagesize)
rstemp.absolutepage=mypage
howmanyrecs=0
howmanyfields=rstemp.fields.count -1
'<p align="center">aaaa</p>
response.write "<div align=""center"">Page " & mypage & " of " & maxpages & "</div><br>"
response.write "<table border='0' align=""center""><tr>"
''Aqui exibe-se o nome das colunas ( pode ser retirado )
'FOR i=0 to howmanyfields
'response.write "<td><b>" & rstemp(i).name & "</b></td>"
'NEXT
response.write "<tr bgcolor=""#0000FF"">" 
response.write "<td width=""26%""><div align=""center""><strong><font color=""#FFFFFF"" size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Col1</font></strong></div></td>"
response.write "<td width=""74%""><div align=""center""><strong><font color=""#FFFFFF"" size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Col2</font></strong></div></td>"
response.write "</tr>"
' Loop dos dados
DO UNTIL rstemp.eof OR howmanyrecs>=maxrecs
response.write "<tr>"
FOR i = 0 to howmanyfields
fieldvalue=rstemp(i)
If isnull(fieldvalue) THEN
fieldvalue="n/a"
END IF
If trim(fieldvalue)="" THEN
fieldvalue=" "
END IF
response.write "<td valign='top'>"
response.write fieldvalue
response.write "</td>"
next
response.write "</tr>"
rstemp.movenext
howmanyrecs=howmanyrecs+1
LOOP
response.write "</table><p>"
' close, destroy
rstemp.close
set rstemp=nothing
' Now make the page _ of _ hyperlinks
Call PageNavBar
sub PageNavBar()
pad=""
scriptname=request.servervariables("script_name")
response.write "<table rows='1' cols='1' width='97%'><tr>"
response.write "<td>"
response.write "<font size='2' color='black' face='Verdana, Arial,Helvetica, sans-serif'>"
if (mypage mod 10) = 0 then
counterstart = mypage - 9
else
counterstart = mypage - (mypage mod 10) + 1
end if
counterend = counterstart + 9
if counterend > maxpages then counterend = maxpages
if counterstart <> 1 then
ref="<a href='" & scriptname
ref=ref & "?whichpage=" & 1
ref=ref & "&pagesize=" & mypagesize
ref=ref & "&sqlQuery=" & server.URLencode(mySQL)
ref=ref & "'>First</a> : "
Response.Write ref

ref="<a href='" & scriptname
ref=ref & "?whichpage=" & (counterstart - 1)
ref=ref & "&pagesize=" & mypagesize
ref=ref & "&sqlQuery=" & server.URLencode(mySQL)
ref=ref & "'>Previous</a> "
Response.Write ref
end if
Response.Write "["
for counter=counterstart to counterend
If counter>=10 then
pad=""
end if
if cstr(counter) <> mypage then
ref="<a href='" & scriptname
ref=ref & "?whichpage=" & counter
ref=ref & "&pagesize=" & mypagesize
ref=ref & "&sqlQuery=" & server.URLencode(mySQL)
ref=ref & "'>" & pad & counter & "</a>"
else
ref="<b>" & pad & counter & "</b>"
end if
response.write ref
if counter <> counterend then response.write " "
next
Response.Write "]"
if counterend <> maxpages then
ref=" <a href='" & scriptname
ref=ref & "?whichpage=" & (counterend + 1)
ref=ref & "&pagesize=" & mypagesize
ref=ref & "&sqlQuery=" & server.URLencode(mySQL)
ref=ref & "'>Next</a>"
Response.Write ref

ref=" : <a href='" & scriptname
ref=ref & "?whichpage=" & maxpages
ref=ref & "&pagesize=" & mypagesize
ref=ref & "&sqlQuery=" & server.URLencode(mySQL)
ref=ref & "'>Last</a>"
Response.Write ref
end if
response.write "<br></font>"
response.write "</td>"
response.write "</table>"
end sub
else
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<p>"
response.write "<font size='2' color='black' face='Verdana, Arial,Helvetica, sans-serif'>"
response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
response.write "Não possui registro a ser impresso"
response.write "<br></font>"
end if

%>
<table width="75%" border="0" align="center">
  <tr bgcolor="#0000FF"> 
    <td width="26%"><div align="center"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Col1</font></strong></div></td>
    <td width="74%"><div align="center"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Col2</font></strong></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<div align="center">aaaa </div>
<p>&nbsp;</p>
<p align="left">aaaa</p>
</body></html>

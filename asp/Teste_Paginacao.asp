<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html><head>
<TITLE>dbtablepaged.asp</TITLE>
</head><body bgcolor="#FFFFFF">
<!--#include file="Adovbs.inc" -->
<% 
'connectme="DSN=Student;uid=student;pwd=magic"
sqltemp="select * from publishers"

'set conn_db = Server.CreateObject("ADODB.Connection")
'conn_db.Open Session("Conn_String_Cogest_Gravacao")

sqltemp = "Select * from transacao"
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
response.Write(mySQL)
response.Write("  -   ")
response.Write(Conn_String_Cogest_Gravacao)
rstemp.open mySQL,Session("Conn_String_Cogest_Gravacao")
rstemp.movefirst
rstemp.pagesize=mypagesize
maxpages=cint(rstemp.pagecount)
maxrecs=cint(rstemp.pagesize)
rstemp.absolutepage=mypage
howmanyrecs=0
howmanyfields=rstemp.fields.count -1
response.write "Page " & mypage & " of " & maxpages & "<br>"
response.write "<table border='1'><tr>"
'Aqui exibe-se o nome das colunas ( pode ser retirado )
FOR i=0 to howmanyfields
response.write "<td><b>" & rstemp(i).name & "</b></td>"
NEXT
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
%>
=======
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html><head>
<TITLE>dbtablepaged.asp</TITLE>
</head><body bgcolor="#FFFFFF">
<!--#include file="Adovbs.inc" -->
<% 
'connectme="DSN=Student;uid=student;pwd=magic"
sqltemp="select * from publishers"

'set conn_db = Server.CreateObject("ADODB.Connection")
'conn_db.Open Session("Conn_String_Cogest_Gravacao")

sqltemp = "Select * from transacao"
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
response.Write(mySQL)
response.Write("  -   ")
response.Write(Conn_String_Cogest_Gravacao)
rstemp.open mySQL,Session("Conn_String_Cogest_Gravacao")
rstemp.movefirst
rstemp.pagesize=mypagesize
maxpages=cint(rstemp.pagecount)
maxrecs=cint(rstemp.pagesize)
rstemp.absolutepage=mypage
howmanyrecs=0
howmanyfields=rstemp.fields.count -1
response.write "Page " & mypage & " of " & maxpages & "<br>"
response.write "<table border='1'><tr>"
'Aqui exibe-se o nome das colunas ( pode ser retirado )
FOR i=0 to howmanyfields
response.write "<td><b>" & rstemp(i).name & "</b></td>"
NEXT
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
%>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
</body></html>
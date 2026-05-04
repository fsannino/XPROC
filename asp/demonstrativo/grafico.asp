<!--#include file="conn_consulta.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("banco.mdb")
db.CursorLocation = 3

set db2 = Server.CreateObject("ADODB.Connection")
db2.Open Session("Conn_String_Cogest_Gravacao")
db2.CursorLocation = 3

		  dim cor(15,2)
		  
		  cor(1,1)="AAAA11"
		  cor(1,2)="AAAA88"

		  cor(2,1)="11AA11"
		  cor(2,2)="11AA88"

		  cor(3,1)="1111BB"
		  cor(3,2)="1188BB"

		  cor(4,1)="AA44BB"
		  cor(4,2)="AA44DD"

		  cor(5,1)="FFCC55"
		  cor(5,2)="FFDD55"

		  cor(6,1)="55D1F3"
		  cor(6,2)="55D1F9"

		  cor(7,1)="DD4411"
		  cor(7,2)="DD7711"

		  cor(8,1)="B1C2D3"
		  cor(8,2)="B1C2D9"

		  cor(9,1)="B1B2B3"
		  cor(9,2)="B1B5B3"

		  cor(10,1)="AB1223"
		  cor(10,2)="AB1201"

		  cor(11,1)="D1C5D3"
		  cor(11,2)="D1C5D3"

		  cor(12,1)="11AA11"
		  cor(12,2)="11AA88"

		  cor(13,1)="1111BB"
		  cor(13,2)="1188BB"

		  cor(14,1)="AA44BB"
		  cor(14,2)="AA44DD"

		  cor(15,1)="FFCC55"
		  cor(15,2)="FFDD55"
		  


set fonte = db2.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

select case request("Abrag")

case "6,8,9"
	tit = "TODAS"
case "6,9"
	tit = "PETROBRAS"
case "8,9"
	tit = "REFAP"
end select

%>

<HTML>
<HEAD>
<TITLE>Demonstrativo de Cursos - Gráficos</TITLE>
</HEAD>
<BODY>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#000099">Acompanhamento 
  de Material Did&aacute;tico</font><font color="#000099"> - Exibi&ccedil;&atilde;o 
  de Gr&aacute;ficos por Mega-Processo</font></b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">Abrang&ecirc;ncia 
  </font><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">: 
  <%=tit%></font></b></p>
<br>

<table width="86%" border="0" height="165">
<%
ssql=""

ssql="SELECT demonstrativo.demo_status AS ESTADO, Count(demonstrativo.demo_cd_curso) AS CONTA "
ssql=ssql+ "FROM demonstrativo INNER JOIN Cursos ON demonstrativo.demo_cd_curso = Cursos.CURS_CD_CURSO "
ssql=ssql+ "WHERE Cursos.ONDA_CD_ONDA In (" & request("Abrag") & ") "
ssql=ssql+ "GROUP BY demonstrativo.demo_status"

'ssql="SELECT demo_status AS ESTADO, Count(demo_cd_curso) AS CONTA "
'ssql=ssql+ "FROM demonstrativo "
'ssql=ssql+ "GROUP BY demo_status"

set rs = db.execute(ssql)
%>
  <tr> 
    <td width="49%" height="198" valign="top"> 
      <table width="96%" border="0">
        <tr valign="top"> 
          <td colspan="2" height="34"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><B><font color="#990000">CURSOS 
            EM ATRASO</font></B></font></td>
        </tr>
		<%
		conta = 0
		
		  do until fonte.eof=true
		  
		  if fonte("MEPR_CD_MEGA_PROCESSO")<>15 then
		  	set temp = db.execute("SELECT * FROM ATRASADOS WHERE MEGA='" & FONTE("MEPR_TX_ABREVIA_CURSO") & "'")
		  else
		  	set temp = db.execute("SELECT * FROM ATRASADOS WHERE MEGA LIKE '" & LEFT(FONTE("MEPR_TX_ABREVIA_CURSO"),2) & "%'")		  
		  end if
		  
		  if temp.eof=false then
		  %>
        <tr bgcolor="#CCCCCC"> 
          <td height="20" width="77%"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=fonte("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
          <td height="20" bgcolor="#FFFFFF" width="23%"> 
            <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><%=TEMP.recordcount%></b></font></div>
		  </td>
			</tr>
		  <%
		  end if
		  fonte.movenext
		  loop
		  
		  fonte.movefirst
		  
		  set temp = nothing
		  %>
		  
      </table>
    </td>
    <td width="51%" height="198"> 
      <div align="center"><applet id=Applet1 style="WIDTH: 450px; HEIGHT: 136px" name=X_PieChart code=X_PieChart.class width=450 height=136>
          <param name="image" value="">
          <param name="forecolor" value="330000">
          <param name="activecolor" value="BBBB00">
          <param name="backcolor" value="FFFFEE">
          <param name="activecolor_shadow" value="DDDD00">
          <param name="distance" value=5>
		  
		<%
		  set fonte = db2.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
		
		  conta = 0
		  atual = 0
		
		  do until fonte.eof=true
		  
		  if fonte("MEPR_CD_MEGA_PROCESSO")<>15 then
		  	set temp = db.execute("SELECT * FROM ATRASADOS WHERE MEGA='" & FONTE("MEPR_TX_ABREVIA_CURSO") & "'")
		  else
		  	set temp = db.execute("SELECT * FROM ATRASADOS WHERE MEGA LIKE '" & LEFT(FONTE("MEPR_TX_ABREVIA_CURSO"),2) & "%'")
		  end if
		  
		  if conta > 11 then
		  	conta = conta - 11
		  end if
		
		  if temp.recordcount>0 then
		  %>
		  <param name="p<%=atual%>" value="<%=fonte("MEPR_TX_DESC_MEGA_PROCESSO")%>#<%=TEMP.RecordCount%>#<%=cor(atual+1,1)%>#<%=cor(atual+1,2)%>###">        
		  <%
  		  atual = atual + 1
		  end if
		  fonte.movenext
		  conta = conta + 1 
		  loop
		  
		  fonte.movefirst
		  %>		  
		  
		  <param name="showratios" value="0">
          <param name="deep" value=3>
          <param name="font" value="Arial#plain#9">
          <param name="title" value="">
        </applet></div>
    </td>
  </tr>
  </table>

<hr>

<table width="72%" border="0" height="165">

<%

ssql=""
ssql="SELECT demonstrativo.demo_status AS ESTADO, Count(demonstrativo.demo_cd_curso) AS CONTA "
ssql=ssql+ "FROM demonstrativo INNER JOIN Cursos ON demonstrativo.demo_cd_curso = Cursos.CURS_CD_CURSO "
ssql=ssql+ "WHERE Cursos.ONDA_CD_ONDA In (" & request("Abrag") & ") "
ssql=ssql+ "GROUP BY demonstrativo.demo_status"

'ssql="SELECT demo_status AS ESTADO, Count(demo_cd_curso) AS CONTA "
'ssql=ssql+ "FROM demonstrativo "
'ssql=ssql+ "GROUP BY demo_status"

set rs = db.execute(ssql)
%>

  <tr> 
    <td width="49%" height="198" valign="top"> 
      <table width="88%" border="0">
        <tr valign="top"> 
          <td colspan="2" height="34"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Mega-Processo 
            : <B>TODOS</B></font></td>
        </tr>
		<%
		  do until rs.eof=true
		  %>
        <tr bgcolor="#CCCCCC"> 
          <td height="20" width="77%"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("estado")%></font></td>
          <td height="20" bgcolor="#FFFFFF" width="23%"> 
            <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><%=rs("conta")%></b></font></div>
		  </td>
			</tr>
		  <%
		  rs.movenext
		  loop
		  %>
		  
      </table>
    </td>
    <td width="51%" height="198"> 
      <div align="center"><applet id=Applet1 style="WIDTH: 450px; HEIGHT: 136px" name=X_PieChart code=X_PieChart.class width=450 height=136>
          <param name="image" value="">
          <param name="forecolor" value="330000">
          <param name="activecolor" value="BBBB00">
          <param name="backcolor" value="FFFFEE">
          <param name="activecolor_shadow" value="DDDD00">
          <param name="distance" value=5>
          
		  <%
		  rs.movefirst
		  
		  conta=0
		  
		  do until rs.eof=true
		  %>
		  <param name="p<%=conta%>" value="<%=rs("ESTADO")%>#<%=rs("CONTA")%>#<%=cor(conta+1,1)%>#<%=cor(conta+1,2)%>###">
          <%
		  rs.movenext
		  conta = conta + 1
		  LOOP
		  %>
		  
		  <param name="showratios" value="0">
          <param name="deep" value=3>
          <param name="font" value="Arial#plain#9">
          <param name="title" value="">
        </applet></div>
    </td>
  </tr>
</table>


<%
do until fonte.eof=true

if fonte("MEPR_CD_MEGA_PROCESSO")=15 then
	pre_curso = left(fonte("MEPR_TX_ABREVIA_CURSO"),2)
else
	pre_curso = left(fonte("MEPR_TX_ABREVIA_CURSO"),3)
end if

ssql=""
ssql="SELECT demonstrativo.demo_status AS ESTADO, Count(demonstrativo.demo_cd_curso) AS CONTA "
ssql=ssql+ "FROM demonstrativo INNER JOIN Cursos ON demonstrativo.demo_cd_curso = Cursos.CURS_CD_CURSO "
ssql=ssql+ "WHERE (Cursos.ONDA_CD_ONDA In (" & request("Abrag") & ")) and (demonstrativo.demo_cd_curso Like '" & pre_curso & "%')"
ssql=ssql+ "GROUP BY demonstrativo.demo_status"

'ssql="SELECT demo_status AS ESTADO, Count(demo_cd_curso) AS CONTA "
'ssql=ssql+ "FROM demonstrativo WHERE demo_cd_curso LIKE '" & pre_curso & "%'"
'ssql=ssql+ "GROUP BY demo_status"

set rs = db.execute(ssql)

if rs.eof=false then
%>

<HR>
<table width="72%" border="0" height="165">
  <tr> 
    <td width="49%" height="198" valign="top"> 
      <table width="88%" border="0">
        <tr valign="top"> 
          <td colspan="2" height="34"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Mega-Processo 
            : <B><%=fonte("MEPR_TX_DESC_MEGA_PROCESSO")%></B></font></td>
        </tr>
		<%
		  do until rs.eof=true
		  %>
        <tr bgcolor="#CCCCCC"> 
          <td height="20" width="77%"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("estado")%></font></td>
          <td height="20" bgcolor="#FFFFFF" width="23%"> 
            <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><%=rs("conta")%></b></font></div>
		  </td>
			</tr>
		  <%
		  rs.movenext
		  loop
		  
		  %>
      </table>
    </td>
    <td width="51%" height="198"> 
      <div align="center"><applet id=Applet1 style="WIDTH: 450px; HEIGHT: 136px" name=X_PieChart code=X_PieChart.class width=450 height=136>
          <param name="image" value="">
          <param name="forecolor" value="330000">
          <param name="activecolor" value="BBBB00">
          <param name="backcolor" value="FFFFEE">
          <param name="activecolor_shadow" value="DDDD00">
          <param name="distance" value=5>
		  
		  <%
		  rs.movefirst
		  
		  conta=0
		  
		  do until rs.eof=true
		  %>
		  <param name="p<%=conta%>" value="<%=rs("ESTADO")%>#<%=rs("CONTA")%>#<%=cor(conta+1,1)%>#<%=cor(conta+1,2)%>###">
          <%
		  rs.movenext
		  conta = conta + 1
		  LOOP
		  %>		  
          

          
		  <param name="showratios" value="0">
          <param name="deep" value=3>
          <param name="font" value="Arial#plain#9">
          <param name="title" value="">
        </applet></div>
    </td>
  </tr>
</table>
<%
	end if
	fonte.movenext
loop
%>


</BODY>
</HTML>
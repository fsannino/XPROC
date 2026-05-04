<%
if request("excel")=1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<!--#include file="../../conn_consulta.asp" -->
<html>
<%
opti=request("op")

set fs=server.CreateObject("Scripting.FileSystemObject")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

Server.ScriptTimeout=99999999

SSQL=""
SSQL="SELECT DISTINCT dbo.CLI.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO "
SSQL=SSQL+"FROM dbo.CLI INNER JOIN "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.CLI.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

'response.write ssql

set rs=db.execute(ssql)
%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="80%" border="0">
  <tr> 
    <td width="55%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
        POR NOME  - COORDENADORES </strong></font></div></td>
    <td width="45%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<table width="98%" border="0" height="97">
  <tr bgcolor="#CCCCCC"> 
    <%if request("excel")=0 then%>
    <td width="19%" height="40" bgcolor="#FFFFFF">&nbsp;</td>
    <%END IF%>
    <td width="19%" height="40"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></td>
    <td width="9%" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Lota&ccedil;&atilde;o</font></strong></div></td>
    <td width="6%" bgcolor="#CCCCCC" height="40"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="6%" bgcolor="#CCCCCC" height="40"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="7%" bgcolor="#CCCCCC" height="40"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
    <td width="26%" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">&Oacute;rgao</font></strong></div></td>
  </tr>
  <%
  do until rs.eof=true
  
  FOTO="http://s600031.serinf.petrobras.com.br/prochn/fotos/" & RIGHT("000000" & RS("MATRICULA"),8) & ".jpg"
  
  if cor="white" then
  	cor="#DCDCED"
  else
  	cor="white"
  end if
%>
  <tr valign="top" bgcolor="<%=cor%>"> 
    <%if request("excel")=1 then%>
    <td height="49"> <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre></td>
    <%else%>
	<td height="49" align="center"> 
    <p align="center">
    <font face="Verdana" size="1" color="#0000FF">
    <img border="0" src="<%=FOTO%>" width="60" height="60" alt="Sem Foto" title=""></font></td>
	<td height="49"> <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre></td>
    <%
	end if
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
	if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
	ON ERROR RESUME NEXT
    %>
	<td height="49"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=temp("LOTACAO")%></font></pre>
      </div></td>
    <td height="49"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre>
      </div></td>
    <td height="49"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
	<td height="49"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("RAMAL"))%></font></pre>
      </div></td>
    <%
    	ssql=""
    	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR AS APOIO FROM CLI_ORGAO WHERE USMA_CD_USUARIO ='" & UCASE(rs("chave")) & "' ORDER BY ORME_CD_ORG_MENOR"
    	
		set rs_orgao=db.execute(ssql)    
    
    	orgao=""
    	do until rs_orgao.eof=true
    		set temp2=db.execute("SELECT ORME_SG_ORG_MENOR AS APOIADO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs_orgao("APOIO") & "'")
			if temp2.eof=true then
			    set temp2=db.execute("SELECT AGLU_SG_AGLUTINADO AS APOIADO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs_orgao("APOIO") & "'")
			end if
			orgao=orgao  & "," & trim(temp2("APOIADO"))
    		rs_orgao.movenext
    	loop
    	
    	if len(orgao)>1 then
			orgao=right(orgao,len(orgao)-1)
		end if
    %>
    <td height="49"> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=orgao%></font></div></td>
  </tr>
  <%
  cli=cli+1
  rs.movenext
  loop
  %>
</table>
<p><font face="Arial, Helvetica, sans-serif" size="2"><b>Total de Coordenadores Locais de Implantação</b> : <%=cli%></font></p>
<p>&nbsp;</p>
</body>
</html>
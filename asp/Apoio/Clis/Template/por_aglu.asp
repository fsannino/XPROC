<!--#include file="../../conn_consulta.asp" -->
<html>
<%
opti=request("op")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

apoio=request("apoio")

dim orgaos(4)
dim codigos(4)

orgaos(1) = "ABAST"
orgaos(2) = "E&P"
orgaos(3) = "REFAP S/A"
orgaos(4) = "OUTROS"

codigos(1)= "88"
codigos(2)= "87"
codigos(3)= "888860000000000"
codigos(4)= "999"

ssql=""
ssql="SELECT DISTINCT AGLU_CD_AGLUTINADO, AGLU_SG_AGLUTINADO "
ssql=ssql+"FROM ORGAO_AGLUTINADOR "
ssql=ssql+"ORDER BY AGLU_SG_AGLUTINADO "

set fonte=db.execute(ssql)
%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="80%" border="0">
  <tr>
    <td width="55%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
        POR AGLUTINADOR&nbsp; - COORDENADORES </strong></font></div></td>
    <td width="45%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<table width="802" border="0" height="155">
<% 
i = 1

do until i = 5

if trim(apoio)=trim(codigos(i)) then
	seleciona=1
	figura="baixo"
else
	seleciona=0
	figura="lado"
end if
%>
<tr bgcolor="#FFFFFF"> 
    <td width="19" height="20"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=orgaos(i)%>"><img src="<%=figura%>.jpg" width="17" height="20"></a></font></td>
    <td width="5" height="20"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>

    <td height="20" colspan="5" width="471"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="por_aglu.asp?op=<%=opti%>&apoio=<%=codigos(i)%>#<%=codigos(i)%>"><%=orgaos(i)%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td height="20" width="167">&nbsp;</td>

    <td height="20" width="172">&nbsp;</td>

    <td width="57" height="20"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="6" height="20">&nbsp;</td>
  </tr>
  <%
  if seleciona=1 then
	
	ssql=""
	ssql="SELECT DISTINCT dbo.CLI.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO, "
	ssql=ssql+"dbo.CLI_ORGAO.ORME_CD_ORG_MENOR AS APOIO "
	ssql=ssql+"FROM dbo.CLI INNER JOIN "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO ON dbo.CLI.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	ssql=ssql+"dbo.CLI_ORGAO ON dbo.CLI.USMA_CD_USUARIO = dbo.CLI_ORGAO.USMA_CD_USUARIO "
	
	if apoio = "999" then
		ssql=ssql+"WHERE (dbo.CLI_ORGAO.ORME_CD_ORG_MENOR NOT LIKE '88%') AND (dbo.CLI_ORGAO.ORME_CD_ORG_MENOR NOT LIKE '87%') AND (dbo.CLI_ORGAO.ORME_CD_ORG_MENOR NOT LIKE '882400300000000%') "
	else
		ssql=ssql+"WHERE (dbo.CLI_ORGAO.ORME_CD_ORG_MENOR LIKE '" & apoio & "%') "
	end if

	ssql=ssql+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"
	
	set rs=db.execute(ssql)
  %>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2" bgcolor="#FFFFFF" width="28" height="40"><font color="#FFFFFF">&nbsp;</font></td>
    <td width="94" height="40"></td>
    <td width="137" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></div></td>
    <td width="74" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="88" bgcolor="#CCCCCC" height="40"><div align="left"><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif"><strong>Lota&ccedil;&atilde;o</strong></font></div></td>
    <td width="62" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="167" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
    <td width="172" bgcolor="#CCCCCC" height="40"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">&Oacute;rgao</font></strong></td>
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
  <tr valign="middle" bgcolor="<%=cor%>"> 
    <td colspan="2" bgcolor="#FFFFFF" width="28" height="57"><font color="#FFFFFF">&nbsp;</font></td>
    <td height="57" valign="top" width="94" align="center"> 
    <p align="center">
    <font face="Verdana" size="1" color="#0000FF">
    <img border="0" src="<%=FOTO%>" width="60" height="60" alt="Sem Foto" title=""></font></td>
    <td height="57" valign="top" width="137"> <div align="left"> 
        <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre>
      </div></td>
    <td valign="top" width="74" height="57"> <div align="left"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre>
      </div></td>
    <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
    if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
	ON ERROR RESUME NEXT    
	%>
    <td valign="top" width="88" height="57"><div align="left"><font size="1" face="Arial, Helvetica, sans-serif"><strong><%=temp("lotacao")%></strong></font></div></td>
    <td valign="top" width="62" height="57"> <div align="left"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
    <td valign="top" width="167" height="57"> <div align="left"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("RAMAL")%></font></pre>
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
        <td height="49" valign="top" width="172"> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=orgao%></font></div></td>
  </tr>
    <%
  tem=tem+1
  rs.movenext
  loop
  end if
  if tem<>0 then
  %>
  <tr valign="middle" bgcolor="#FFFFFF"> 
    <td colspan="2" width="28" height="22">&nbsp;</td>
    <td height="22" align="center" valign="middle" bgcolor="#CCCCCC" width="94">&nbsp;</td>
    <td height="22" align="center" valign="middle" bgcolor="#CCCCCC" width="137"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
      de Registros : </strong><font size="2"><%=tem%></font></font></td>
    <td width="74" height="22">&nbsp;</td>
    <td width="88" height="22">&nbsp;</td>
    <td width="62" height="22">&nbsp;</td>
    <td valign="top" width="167" height="57"> &nbsp;</td>
    <td width="172" height="22">&nbsp;</td>
  </tr>
  <%
  end if
  tem=0
  seleciona=0
  i = i + 1
  loop
  %>
</table>
</body>
</html>
<!--#include file="../../conn_consulta.asp" -->
<html>
<%
opti=request("op")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

apoio=request("apoio")

ssql=""
ssql="SELECT DISTINCT CLI_ORGAO.ORME_CD_ORG_MENOR, ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO, ORGAO_MENOR.ORME_SG_ORG_MENOR "
ssql=ssql+"FROM CLI_ORGAO LEFT JOIN ORGAO_AGLUTINADOR ON "
ssql=ssql+"CLI_ORGAO.ORME_CD_ORG_MENOR = ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO LEFT JOIN ORGAO_MENOR ON "
ssql=ssql+"CLI_ORGAO.ORME_CD_ORG_MENOR = ORGAO_MENOR.ORME_CD_ORG_MENOR "
ssql=ssql+"ORDER BY ORGAO_MENOR.ORME_SG_ORG_MENOR , ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO "

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
        POR ORG&Atilde;O - COORDENADORES </strong></font></div></td>
    <td width="45%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<table width="931" border="0" height="159">
<%
do until fonte.eof=true
if trim(apoio)=trim(fonte("ORME_CD_ORG_MENOR")) then
	seleciona=1
	figura="baixo"
else
	seleciona=0
	figura="lado"
end if
%>
<tr bgcolor="#FFFFFF"> 
    <td width="19" height="24"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte("ORME_CD_ORG_MENOR")%>"><img src="<%=figura%>.jpg" width="17" height="20"></a></font></td>
    <td width="7" height="24"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <%IF LEN(fonte("ORME_CD_ORG_MENOR"))=2 THEN%>
    <td height="24" colspan="5" width="786"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_apoio.asp?op=<%=opti%>&apoio=<%=fonte("ORME_CD_ORG_MENOR")%>#<%=fonte("ORME_CD_ORG_MENOR")%>"><%=rtrim(fonte("AGLU_SG_AGLUTINADO"))%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td height="24" width="336">&nbsp;</td>
    <%ELSE%>
    <td height="24" width="126"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_apoio.asp?op=<%=opti%>&apoio=<%=fonte("ORME_CD_ORG_MENOR")%>#<%=fonte("ORME_CD_ORG_MENOR")%>"><%=rtrim(fonte("ORME_SG_ORG_MENOR"))%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <%END IF%>
    <td width="174" height="24"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="18" height="24">&nbsp;</td>
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
	ssql=ssql+"WHERE (dbo.CLI_ORGAO.ORME_CD_ORG_MENOR = '" & apoio & "') "
	ssql=ssql+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"
	
	set rs=db.execute(ssql)
  %>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2" bgcolor="#FFFFFF" width="30" height="40"><font color="#FFFFFF">&nbsp;</font></td>
    <td width="131" height="40"></td>
    <td width="241" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></div></td>
    <td width="90" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="133" bgcolor="#CCCCCC" height="40"><div align="left"><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif"><strong>Lota&ccedil;&atilde;o</strong></font></div></td>
    <td width="106" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="21" bgcolor="#CCCCCC" height="40"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
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
    <td colspan="2" bgcolor="#FFFFFF" width="30" height="57"><font color="#FFFFFF">&nbsp;</font></td>
    <td height="57" valign="middle" width="131" align="center"> 
    <p align="center">
    <font face="Verdana" size="1" color="#0000FF">
    <img border="0" src="<%=FOTO%>" width="60" height="60" alt="Sem Foto" title=""></font></td>
    <td height="57" valign="top" width="241"> <div align="left"> 
        <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre>
      </div></td>
    <td valign="top" width="90" height="57"> <div align="left"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre>
      </div></td>
    <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
    if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
	ON ERROR RESUME NEXT    
	%>
    <td valign="top" width="133" height="57"><div align="left"><font size="1" face="Arial, Helvetica, sans-serif"><strong><%=temp("lotacao")%></strong></font></div></td>
    <td valign="top" width="106" height="57"> <div align="left"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
    <td valign="top" width="21" height="57"> <div align="left"> 
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
    <%
  tem=tem+1
  rs.movenext
  loop
  end if
  if tem<>0 then
  %>
  <tr valign="middle" bgcolor="#FFFFFF"> 
    <td colspan="2" width="30" height="22">&nbsp;</td>
    <td height="22" align="center" valign="middle" bgcolor="#CCCCCC" width="131">&nbsp;</td>
    <td height="22" align="center" valign="middle" bgcolor="#CCCCCC" width="241"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
      de Registros : </strong><font size="2"><%=tem%></font></font></td>
    <td width="90" height="22">&nbsp;</td>
    <td width="133" height="22">&nbsp;</td>
    <td width="106" height="22">&nbsp;</td>
    <td width="21" height="22">&nbsp;</td>
  </tr>
  <%
  end if
  tem=0
  seleciona=0
  fonte.movenext
  loop
  %>
</table>
</body>
</html>
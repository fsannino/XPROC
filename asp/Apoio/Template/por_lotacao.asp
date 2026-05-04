<!--#include file="../conn_consulta.asp" -->
<html>
<%
server.ScriptTimeout=99999999

opti=request("op")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

lotacao=request("lotacao")

ssql=""
ssql="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR "
ssql=ssql+"FROM dbo.ORGAO_MENOR INNER JOIN "
ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1) "
ssql=ssql+"ORDER BY dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR "

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
        POR LOTA&Ccedil;&Atilde;O - APOIADOR LOCAL</strong></font></div></td>
    <td width="45%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<table width="965" border="0">
  <%
do until fonte.eof=true
if trim(lotacao)=trim(fonte("ORME_CD_ORG_MENOR")) then
	seleciona=1
	figura="baixo"
else
	seleciona=0
	figura="lado"
end if
%>
  <tr bgcolor="#FFFFFF"> 
    <td width="20"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a name="<%=fonte("ORME_CD_ORG_MENOR")%>"><img src="<%=figura%>.jpg" width="17" height="20"></a></font></td>
    <td width="6"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td height="28" colspan="4" width="423"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="por_lotacao.asp?op=<%=opti%>&lotacao=<%=fonte("ORME_CD_ORG_MENOR")%>#<%=fonte("ORME_CD_ORG_MENOR")%>"><%=rtrim(fonte("ORME_SG_ORG_MENOR"))%></a></font></strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="258"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="226"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
  <%
  if seleciona=1 then
	SSQL=""
	SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR AS LOTACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
	SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1) AND (dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR='" & lotacao & "')"
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

	set rs=db.execute(ssql)
	
  %>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2" bgcolor="#FFFFFF" width="32"><font color="#FFFFFF">&nbsp;</font></td>
    <td width="195" height="44"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></td>
    <td width="70" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="41" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="99" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
    <td width="258" bgcolor="#CCCCCC"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Assunto</font></strong></div></td>
    <td width="226" bgcolor="#CCCCCC"><div align="left"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">&Oacute;rgao 
        Apoiado</font></strong></div></td>
  </tr>
  <%
  do until rs.eof=true
  if cor="white" then
  	cor="#DCDCED"
  else
  	cor="white"
  end if
  %>
  <tr valign="middle" bgcolor="<%=cor%>"> 
    <td colspan="2" bgcolor="#FFFFFF" width="32"><font color="#FFFFFF">&nbsp;</font></td>
    <td height="26" valign="top" width="195"> <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre></td>
    <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
    if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
    %>
    <td valign="top" width="70"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre>
      </div></td>
    <td valign="top" width="41"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
    <td valign="top" width="99"> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("RAMAL")%></font></pre>
      </div></td>
    <%
		ssql=""
		ssql="SELECT distinct dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS NOME "
		ssql=ssql+"FROM dbo.SUB_MODULO INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON "
		ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
		ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = '" & rs("chave") & "') AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO=1)"
		ssql=ssql+" ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"
		
		set rs_modulo=db.execute(ssql)
		
		modulo=""
		do until rs_modulo.eof=true
			modulo=modulo & "," & rs_modulo("NOME")		
			rs_modulo.movenext
		loop
		
		if len(modulo)>1 then
			modulo=right(modulo,len(modulo)-1)
		end if
    %>
    <td valign="top" width="258"> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=modulo%></font></div></td>
    <%
    	ssql=""
    	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR AS APOIO FROM APOIO_LOCAL_ORGAO WHERE (USMA_CD_USUARIO ='" & UCASE(rs("chave")) & "') AND (APLO_NR_ATRIBUICAO=1) AND SUBSTRING(ORME_CD_ORG_MENOR,11,5) = '00000' ORDER BY ORME_CD_ORG_MENOR"
    	
		set rs_orgao=db.execute(ssql)    
    
    	orgao=""
    	do until rs_orgao.eof=true
    		set temp2=db.execute("SELECT ORME_SG_ORG_MENOR AS APOIADO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs_orgao("APOIO") & "'")
			if temp2.eof=true then
			    set temp2=db.execute("SELECT AGLU_SG_AGLUTINADO AS APOIADO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs_orgao("APOIO") & "'")
			end if
			orgao=orgao  & ", " & trim(temp2("APOIADO"))
    		rs_orgao.movenext
    	loop
    	
    	if len(orgao)>1 then
			orgao=right(orgao,len(orgao)-1)
		end if
    %>
    <td valign="top" width="226"> 
      <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=orgao%></A></div></td>
  </tr>
  <%
  tem=tem+1
  rs.movenext
  loop
  end if
  if tem<>0 then
  %>
  <tr valign="middle" bgcolor="#FFFFFF"> 
    <td colspan="2" width="32">&nbsp;</td>
    <td height="26" align="center" valign="middle" bgcolor="#CCCCCC" width="195"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
      de Registros : </strong><font size="2"><%=tem%></font></font></td>
    <td width="70">&nbsp;</td>
    <td width="41">&nbsp;</td>
    <td width="99">&nbsp;</td>
    <td valign="top" width="258">&nbsp;</td>
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

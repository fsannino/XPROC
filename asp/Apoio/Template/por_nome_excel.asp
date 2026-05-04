<!--#include file="../conn_consulta.asp" -->
<%
Response.Buffer = true
Response.ContentType = "application/vnd.ms-excel"
%>
<html>
<%
opti=request("op")
tipo=request("tipo")

if tipo=1
	quali="Apoiado"
else
	quali="Multiplicado"
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

Server.ScriptTimeout=99999999

SSQL=""
SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO, "
SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO,  dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO "
SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & tipo & ")"
SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

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
<table width="98%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="16%" height="44"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Nome</font></strong></td>
    <td width="7%" bgcolor="#CCCCCC"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Lota&ccedil;&atilde;o</font></strong></div></td>
    <td width="7%" bgcolor="#CCCCCC"><div align="center"><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif"><strong>Momento</strong></font></div></td>
    <td width="16%" height="44"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Atribuição</font></strong></div></td>
    <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Matricula</font></strong></div></td>
    <td width="4%" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Chave</font></strong></div></td>
    <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Ramal</font></strong></div></td>
    <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Onda</font></strong></div></td>
    <td width="11%" bgcolor="#CCCCCC"><div align="left"><strong><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif">Assunto</font></strong></div></td>
    <td width="36%" bgcolor="#CCCCCC"><div align="left"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">&Oacute;rgao <%=quali%></font></strong></div></td>
  </tr>
  <%
  do until rs.eof=true
  if cor="white" then
  	cor="#DCDCED"
  else
  	cor="white"
  end if
%>
  <tr valign="top" bgcolor="<%=cor%>"> 
    <td height="26"> <pre><font color="#0000FF" size="1" face="Arial, Helvetica, sans-serif"><%=rtrim(rs("NOME"))%></font></pre></td>
    <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
	if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
	ON ERROR RESUME NEXT
		set temp3=db.execute("SELECT APLO_NR_MOMENTO AS MOMENTOS FROM APOIO_LOCAL_MULT WHERE USMA_CD_USUARIO='" & UCASE(rs("CHAVE")) & "' AND APLO_NR_SITUACAO=1")	
		
		total=""
		
		do until temp3.eof=true
			total=total & temp3("MOMENTOS")
			temp3.movenext
		loop
		
		m1=0
		m2=0
		
		m1=instr(total, "1")
		m2=instr(total, "2")
		
		momento=""
		
		if m1<>0 and m2<>0 then
			momento="1 e 2"
		else
			if m1=0 and m2<>0 then
				momento="2"
			else
				if m1<>0 and m2=0 then
					momento="1"
				else
					if m1<>0 and m2=0 then
						momento=" - "
					end if
				end if
			end if
		end if
	%>
        <%
	set temp3=db.execute("SELECT APLO_NR_ATRIBUICAO FROM APOIO_LOCAL_MULT WHERE USMA_CD_USUARIO='" & UCASE(rs("CHAVE")) & "' ORDER BY APLO_NR_ATRIBUICAO")
	
	atrib_atual=""
	
	do until temp3.eof=true
		atrib_atual = atrib_atual & temp3("APLO_NR_ATRIBUICAO")
		temp3.movenext
	loop
	
	select case atrib_atual
		case "1"
			atribuicao="APOIADOR LOCAL"
		case "2"
			atribuicao="MULTIPLICADOR"
		case "12"
			atribuicao="APOIADOR LOCAL / MULTIPLICADOR"
	end select
	
	%>
	<td> <div align="center"><pre><font size="1" face="Arial, Helvetica, sans-serif"><%=temp("lotacao")%></font></pre></div></td>
	<td> <div align="center"><pre><font size="1" face="Arial, Helvetica, sans-serif"><%=momento%></font></pre></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=atribuicao%></font></div></td>	
	<td> <div align="center"><pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("MATRICULA")%></font></pre></div></td>
    <td> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=UCASE(rs("CHAVE"))%></font></pre>
      </div></td>
    <td> <div align="center"> 
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=rs("RAMAL")%></font></pre>
      </div></td>
    <td> <div align="center"> 
        <%
      set onda = db.execute("SELECT * FROM APOIO_LOCAL_ONDA WHERE USMA_CD_USUARIO='" & UCASE(rs("CHAVE")) & "' AND APLO_NR_ATRIBUICAO = " & tipo )
      sl_onda=""
            
      do until onda.eof=true
      		if len(onda("ONDA_CD_ONDA"))>1 then
				at_onda=right(onda("ONDA_CD_ONDA"),1)
			else
				at_onda=onda("ONDA_CD_ONDA")
			end if
			
			set rsonda=db.execute("SELECT * FROM ONDA WHERE ONDA_CD_ONDA=" & at_onda)
			
	 		sl_onda=sl_onda & rsonda("ONDA_TX_ABREV_ONDA") & ","   
	      	
	      	onda.movenext
      loop
     
      sl_onda=left(sl_onda,len(sl_onda)-1)
      
      %>
        <pre><font size="1" face="Arial, Helvetica, sans-serif"><%=sl_onda%></font></pre>
      </div></td>
    <%
		ssql=""
		ssql="SELECT distinct dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS NOME "
		ssql=ssql+"FROM dbo.SUB_MODULO INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON "
		ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
		ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = '" & rs("chave") & "') AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & tipo & ")"
		ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"
		
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
    <td> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=modulo%></font></div></td>
    <%
    	ssql=""
    	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR AS APOIO FROM APOIO_LOCAL_ORGAO WHERE (USMA_CD_USUARIO ='" & UCASE(rs("chave")) & "') AND (SUBSTRING(ORME_CD_ORG_MENOR,11,5) = '00000') AND (APLO_NR_ATRIBUICAO = " & tipo & ") ORDER BY ORME_CD_ORG_MENOR"
    	
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
    <td> <div align="left"> <font size="1" face="Arial, Helvetica, sans-serif"><%=orgao%></font></div></td>
  </tr>
  <%
  rs.movenext
  loop
  %>
</table>
</body>
</html>
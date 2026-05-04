<!--#include file="../conn_consulta.asp" -->
<html>
<%
opti=request("op")

tipo=Session("Tipo")
chave=request("chave")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

SSQL=""
SSQL="SELECT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO AS ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_RELACAO_EMPREGO AS VINCULO, dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO AS MOMENTO, "
SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO, dbo.APOIO_LOCAL_MULT.APLO_TX_OBS AS OBSERVACAO "
SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO='" & CHAVE & "' "
SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

set rs=db.execute(ssql)

select case rs("VINCULO")
case "C"
	val_vinc="CONTRATADO"
case "E"
	val_vinc="EMPREGADO"
end select

do until rs.eof=true
	atribui=atribui & rs("atribuicao")
	rs.movenext
loop

rs.movefirst

val_atribui1=""
val_atribui2=""

select case atribui
	case 1
		val_atribui1="APOIADOR LOCAL"
	case 2
		val_atribui1="MULTIPLICADOR"
	case 12
		val_atribui1="APOIADOR LOCAL "
		val_atribui2=" MULTIPLICADOR"	
	case 21
		val_atribui1="APOIADOR LOCAL "
		val_atribui2=" MULTIPLICADOR"	
end select

%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="75%" border="0">
  <tr> 
    <td width="10%"><div align="right"><a href="javascript:history.go(-1)"><img src="volta_f02.gif" width="24" height="24" border="0"></a></div></td>
    <td width="32%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
    <td><strong><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif">Apoiadores 
      Locais e Multiplicadores</font></strong></td>
  </tr>
  <tr> 
    <td><div align="right"><a href="javascript:print()"><img src="../impressao.jpg" width="25" height="24" border="0"></a></div></td>
    <td><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Imprimir</strong></font></td>
    <td width="58%"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
      ( Clique na atribui&ccedil;&atilde;o para editar o registro )</font></strong></td>
  </tr>
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="53%" border="0">
  <tr> 
    <td height="27" bgcolor="#000099"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atribui&ccedil;&atilde;o</strong></font></td>
    <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    <%if opti=1 then%>
	<%IF TIPO=0 OR TIPO=2 THEN%>
	<b><a href="../cad_apoio.asp?tipo=1&valor=<%=rs("chave")%>" target="_top" title="Clique aqui para editar o Apoiador Local"><%=val_atribui1%></a> 
	<%ELSE%>
	<b><%=val_atribui1%></b>
	<%END IF%>
      <%if len(val_atribui2)>0 then%>
      / 
      <%end if%>
      <%IF TIPO=1 OR TIPO=2 THEN%>
	  <B><a href="../cad_apoio.asp?tipo=2&valor=<%=rs("chave")%>" target="_top" title="Clique aqui para editar o Multiplicador"><%=val_atribui2%></a></font></b></td>
	  <%ELSE%>
	  <b><%=val_atribui2%></font></b></td>	  
	  <%END IF%>
	  <%else%>
      <b><%=val_atribui1%></b>
      <%if len(val_atribui2)>0 then%>
      / 
      <%end if%>
	  <b><%=val_atribui2%></font></b></td>	  
	  <%end if%>
  </tr>
  <tr> 
    <td width="16%" height="27" bgcolor="#FFFFFF"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="49%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=VAL_VINC%></strong></font></td>
    <%
	SELECT CASE rs("situacao")
	case 1
	%>
    <td width="17%" bgcolor="#CCCCCC">
<div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>ATIVO</strong></font></div></td>
    <%
	case else
	%>
    <td width="18%" bgcolor="#CCCCCC">
<div align="center"><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>INATIVO</strong></font></div></td>
    <%end select%>
  </tr>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="2" height="165">
  <tr bgcolor="#FFFFFF"> 
    <td height="21" width="73">&nbsp;</td>
    <td width="203" height="21">&nbsp;</td>
    <td width="122" height="21">&nbsp;</td>
    <td width="380" height="21">&nbsp;</td>
  </tr>
  <tr> 
    <td width="73" height="29" bgcolor="#000099"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Matr&iacute;cula</font></strong></td>
    <td width="203" height="29"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("matricula")%></font></td>
    <td width="122" bgcolor="#000099" height="29"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Ramal</font></strong></td>
    <td width="380" height="29"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ramal")%></font></td>
  </tr>
  <tr> 
    <td height="24" bgcolor="#000099" width="73"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nome</font></strong></td>
    <td width="203" height="24"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("nome")%></font></td>
    <td bgcolor="#000099" width="122" height="24"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">M&oacute;dulo</font></strong></td>
    <%
		ssql=""
		ssql="SELECT distinct dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS NOME "
		ssql=ssql+"FROM dbo.SUB_MODULO INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON "
		ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
		ssql=ssql+"dbo.APOIO_LOCAL_MULT ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO "
		ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = '" & rs("chave") & "')"
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
    <td width="380" height="24"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=modulo%></font></td>
  </tr>
  <tr> 
    <td height="34" bgcolor="#000099" width="73"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Lota&ccedil;&atilde;o</font></strong></td>
    <%
    set temp=db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & rs("lotacao") & "'")
    if temp.eof=true then
	    set temp=db.execute("SELECT AGLU_SG_AGLUTINADO AS LOTACAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & rs("lotacao") & "'")
    end if
    %>
    <td width="203" height="34"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=temp("lotacao")%></font></td>
    <td bgcolor="#000099" width="122" height="34"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;os 
      Apoiados</font></strong></td>
    <%
    	ssql=""
    	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR AS APOIO FROM APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO ='" & UCASE(rs("chave")) & "' ORDER BY ORME_CD_ORG_MENOR"
    	
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
    <td width="380" height="75" rowspan="2" align="left" valign="top"><p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=orgao%></font></p></td>
  </tr>
  <tr> 
    <td height="37" bgcolor="#000099" width="73"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Chave</font></strong></td>
    <td width="203" height="37"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=UCASE(rs("chave"))%></font></td>
    <td width="122" height="37"><font size="1">&nbsp;</font></td>
  </tr>
</table>
<table width="71%" border="0">
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#000099"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Onda</font></strong></td>
  </tr>
  <%
	SSQL=""
	SSQL="SELECT dbo.APOIO_LOCAL_ONDA.USMA_CD_USUARIO, dbo.ONDA.ONDA_TX_DESC_ONDA AS ONDA "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_ONDA INNER JOIN "
	SSQL=SSQL+"dbo.ONDA ON dbo.APOIO_LOCAL_ONDA.ONDA_CD_ONDA = dbo.ONDA.ONDA_CD_ONDA "
	SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_ONDA.USMA_CD_USUARIO = '" & rs("chave") & "')"
	SSQL=SSQL+"ORDER BY dbo.ONDA.ONDA_TX_DESC_ONDA"
	
	set rs_onda=db.execute(ssql)
	
		onda=""
		do until rs_onda.eof=true
			onda=onda & "," & rs_onda("ONDA")		
			rs_onda.movenext
		loop
		
		if len(onda)>1 then
			onda=right(onda,len(onda)-1)
		end if
  
  %>
  <tr> 
    <td bgcolor="#FFFFFF"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=onda%></font></td>
  </tr>
</table>
<table width="96%" border="0">
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#000099"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Momento</strong></font></td>
  </tr>
  <%
  SELECT CASE RS("MOMENTO")
  CASE 1
  	VAL1="checked"
	VAL2=""
  CASE 2
  	VAL1=""
	VAL2="checked"
  CASE 12
  	VAL1="checked"
	VAL2="checked"
  case else
  	VAL1=""
	VAL2=""
  END SELECT  
  %>
  <tr> 
    <td bgcolor="#000099"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
      <input name="mom1" type="checkbox" id="mom1" value="checkbox" <%=val1%>>
      </font> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
      Momento 1 - </font><font color="#FFFFFF" size="1" face="Verdana">Completeza; 
      Mapeamentos Treinamento e Perfil; Testes Integrados</font></td>
  </tr>
  <tr> 
    <td bgcolor="#000099"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
      <input name="mom2" type="checkbox" id="mom2" value="checkbox" <%=val2%>>
      Momento 2 - </font><font color="#FFFFFF" size="1" face="Verdana">Partida 
      e Estabilização</font></td>
  </tr>
</table>
<table width="75%" border="0">
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#000099"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es</font></strong></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=RS("OBSERVACAO")%></font></td>
  </tr>
</table>
<p>&nbsp;</p></body>
</html>

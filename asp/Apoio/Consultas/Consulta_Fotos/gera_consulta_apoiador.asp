<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
Server.ScriptTimeOut=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set fso=server.createobject("Scripting.FileSystemObject")

atribui=request("atrib")
if atribui = 1 THEN
	TIT="Apoiadores Locais"
ELSE
	TIT="Multiplicadores"
END IF

org=request("org")

modo=request("modo")

if org=1 then
	ORGAO = "APOIO_LOCAL_ORGAO"
else
	ORGAO = "APOIO_LOCAL_MULT"
end if

orgao1=request("str01")
orgao2=request("str02")
orgao3=request("str03")
orgao4=request("str04")
orgao5=request("str05")

modulo=request("selModulo_")

if len(modulo)<1 then
	modulo=0
end if

if orgao1=0 and orgao2=0 and orgao3=0 and orgao4=0 and orgao5=0 and modulo<>0 then 

	SSQL=""
	SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO AS ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO AS MOMENTO, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR AS LOTACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS MODULO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
	SSQL=SSQL+"dbo.ORGAO_MENOR ON dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND "
	if request("MOMENTO")<>0 THEN
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO = " & request("momento") & ") AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) "
	ELSE
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) "
	END IF
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "	
	
	set fonte = db.execute(ssql)
	
else

	conta = 5

	if modo=1 then
		orgao2 = orgao2 & "00000000"
		orgao3 = orgao3 & "00000"
		orgao4 = orgao4 & "00"
	end if
	
	orgao_final=orgao5
	
	if modo=1 then
	
	if len(orgao5)<15 then
		conta=4
		orgao_final=orgao4
	end if

	if len(orgao4)<15 then
		conta=3
		orgao_final=orgao3
	end if

	if len(orgao3)<15 then
		conta=2
		orgao_final=orgao2
	end if

	if len(orgao2)<15 then
		conta=1
		orgao_final=orgao1
	end if
	
	else

	if len(orgao5)<15 then
		conta=4
		orgao_final=orgao4
	end if

	if len(orgao4)<13 then
		conta=3
		orgao_final=orgao3
	end if

	if len(orgao3)<10 then
		conta=2
		orgao_final=orgao2
	end if

	if len(orgao2)<7 then
		conta=1
		orgao_final=orgao1
	end if
	
	end if
	
	SSQL=""
	SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO AS ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO AS MOMENTO, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR AS LOTACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS MODULO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
	SSQL=SSQL+"dbo.ORGAO_MENOR ON dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	
	if request("MOMENTO")<>0 THEN
		SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_MULT.APLO_NR_MOMENTO = " & request("momento") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & atribui & ")"
	ELSE
		SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & atribui & ")"
	END IF
	
	if modo=1 then
		SSQL=SSQL+"AND (dbo." & ORGAO & ".ORME_CD_ORG_MENOR = '" & orgao_final & "') AND "
	else
		SSQL=SSQL+"AND (dbo." & ORGAO & ".ORME_CD_ORG_MENOR LIKE '" & orgao_final & "%') AND "
	end if
	
	if modulo<>0 then
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) "
	else
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) "
	end if	
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
	
	set fonte=db.execute(ssql)
	
	'response.write ssql

end if

if fonte.eof=true then

	ssql=""
	ssql="SELECT DISTINCT dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA "
	ssql=ssql+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = 000 ) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND "
	ssql=ssql+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) "
	ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR"
	
	set fonte = db.execute(ssql)
	
	achou=0

else
	
	achou=1

end if

'response.write ssql

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Base de Dados de Apoiadores Locais</title>
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
if(window.event.keyCode==121)
{
window.history.go(-1);
return;
}
if(window.event.keyCode==120)
{
window.print();
}
}
</script>

<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">

<form name="frm1" method="POST" action="gera_consulta_apoiador_excel.asp" target="blank">
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
   <table border="0" width="95%">
    <tr>
      <td width="35%" rowspan="3">
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
   <font size="5" color="#000080" face="Tahoma"><b><%=TIT%></b></font></p>
      </td>
      <td width="4%" bordercolor="#31009C" bgcolor="#FFFFFF" align="right">
        <a href="javascript:print()"><img border="0" src="impressao.jpg" width="28" height="30"></a></td>
      <td width="22%" bordercolor="#31009C" bgcolor="#FFFFFF">
        <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" size="2" face="Verdana">Imprimir
        Consulta</font></b></p>
      </td>
    </tr>
    <tr>
      <td width="4%" bordercolor="#31009C" bgcolor="#FFFFFF" align="right">
        <a href="javascript:history.go(-1)"><img border="0" src="volta_f02.gif" width="24" height="24"></a></td>
      <td width="22%" bordercolor="#31009C" bgcolor="#FFFFFF">
        <b><font color="#000080" size="2" face="Verdana"> Voltar para a Tela
        Anterior</font></b></td>
    </tr>
    <tr>
      <td width="4%" bordercolor="#31009C" bgcolor="#FFFFFF" align="right">
        <a target="_blank" href="gera_consulta_apoiador_excel.asp?ssql=<%=replace(ssql,"%","*")%>"><img border="0" src="excel.jpg" width="22" height="21"></a></td>
      <td width="22%" bordercolor="#31009C" bgcolor="#FFFFFF">
        <b><font color="#000080" size="2" face="Verdana">Exportar consulta para
        o Excel</font></b></td>
    </tr>
   </table>
   <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
                <table border="1" width="97%" bordercolor="#C0C0C0" cellspacing="0" cellpadding="2">
                  <%
                  if fonte.eof=false then
                  %>
                  <tr>
                    <td width="20%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">ÓRGÃO
                      APOIADO</font></b></td>
                    <td width="30%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">NOME&nbsp;</font></b></td>
                    <td width="17%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">LOTAÇÃO</font></b></td>
                    <td width="12%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">MÓDULO</font></b></td>
                    <td width="10%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">MOMENTO</font></b></td>
                    <td width="10%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">CHAVE</font></b></td>
                    <td width="10%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">FOTO</font></b></td>
                    <td width="13%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">TELEFONE</font></b></td>
                  </tr>
                  <%
                  end if
                  %>
                  <%
                  chave_ant=0
                  igual=0
                  
                  do until fonte.eof=true
                  chave_atual=fonte("chave")
                  
                  if trim(chave_ant)<>trim(chave_atual) then
                  	igual=igual+1
                  end if
                  
                  if cor="#E1E1E1" then
                  	cor="white"
                  else
                  	cor="#E1E1E1"
                  end if
                  %>
                  <tr>
                    <%
						ORG_APOIO=" "
						SET TEMP=DB.EXECUTE("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("apoio") & "'")
						on error resume next
						ORG_APOIO=UCASE(TEMP("ORME_SG_ORG_MENOR"))
						if err.number<>0 then
							SET TEMP=DB.EXECUTE("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & fonte("apoio") & "'")
							on error resume next
							ORG_APOIO=UCASE(TEMP("AGLU_SG_AGLUTINADO"))
							if err.number<>0 then
								ORG_APOIO=" "
							END IF
						END IF
                    %>
                    <td width="20%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><b><%=ORG_APOIO%></b></font></td>
                    <td width="30%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><b><%=UCASE(fonte("NOME"))%></b></font></td>
                    <td width="17%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=UCASE(fonte("LOTACAO"))%></font></td>
                    <td width="12%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=UCASE(fonte("MODULO"))%></font></td>
                    <%
                    MOMENTO=FONTE("MOMENTO")
                    SELECT CASE MOMENTO
                    CASE 0
                    		MOM_ATUAL=""
                    CASE NULL
                    		MOM_ATUAL=""
                    CASE ""
                    		MOM_ATUAL=""
                    CASE 12
                    		MOM_ATUAL="1 E 2"
                    CASE ELSE
	                    	MOM_ATUAL=FONTE("MOMENTO")
	                 END SELECT
                    %>
                    <td width="10%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=MOM_ATUAL%></font></td>
                    <td width="10%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=UCASE(fonte("CHAVE"))%></font></td>
                    <td width="10%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1">
				      <%
				      	TEM_FOTO=FALSE
				      	CAMINHO = SERVER.MAPPATH("FOTOS\"& fonte("CHAVE") & ".jpg")
				      	TEM_FOTO = FSO.FILEEXISTS(CAMINHO)
				       
				       IF TEM_FOTO=TRUE THEN
							FOTO=fonte("CHAVE")
						ELSE
							FOTO="SEM_FOTO"				      
						END IF
				      %>
				      <img border="1" src="fotos/<%=FOTO%>.jpg" width="60" height="60"></font></td>
                    <td width="13%" align="center" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=UCASE(fonte("RAMAL"))%></font></td>
                  </tr>
                  <%
                  TEM=TEM+1
                  chave_ant=fonte("chave")
                  fonte.movenext
                  loop
                  %>
                  </table>
                  <p>
                  <%
                  if achou=0 then
                  %>
					<p align="center"><font color="#800000"><b>&nbsp;Nenhum Registro Encontrado para a Seleção</b></font></p>
					<%else%>
					<p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#800000">&nbsp;</font><font color="#000080" size="2" face="Verdana">Total de Registros Encontrados :
                    </font> </b><font color="#000080" size="2" face="Verdana"><%=tem%></font></p>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
                    &nbsp;<font color="#000080" size="2" face="Verdana"><b>Total de
                    Apoiadores / Multiplicadores :
 </b>
                    <%=igual%><b>
 </b>
                    </font>
                    
					<%
					end if
					%>
                    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
                    <font color="#000080" size="2" face="Verdana"></font>
</form>
</body>

</html>

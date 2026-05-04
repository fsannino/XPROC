<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"

Server.ScriptTimeOut=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

query = request("ssql")
query=replace(query,"*","%")

set fonte=db.execute(query)

if fonte.eof=true then

	ssql=""
	ssql="SELECT TOP 100 PERCENT dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA "
	ssql=ssql+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = 000 ) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1) AND "
	ssql=ssql+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) "
	ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR"
	
	set fonte = db.execute(ssql)
	
	achou=0

else
	
	achou=1

end if

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

<script language="javascript" src="../js/troca_lista2.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">

<form>
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
                    <td width="20%" align="center"><font face="Verdana" size="1"><b><%=ORG_APOIO%></b></font></td>
                    <td width="30%" align="center"><font face="Verdana" size="1"><b><%=UCASE(fonte("NOME"))%></b></font></td>
                    <td width="17%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("LOTACAO"))%></font></td>
                    <td width="12%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("MODULO"))%></font></td>
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
                    <td width="10%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("CHAVE"))%></font></td>
                    <td width="13%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("RAMAL"))%></font></td>
                  </tr>
                  <%
                  TEM=TEM+1
                  chave_ant=fonte("chave")

                  fonte.movenext
                  loop
                  %>
                  </table>                  <p>
                  <%
                  if achou=0 then
                  %>
					<p align="center"><font color="#800000"><b>&nbsp;Nenhum Registro Encontrado para a Seleção</b></font></p>
					<%else%>
					<p align="left" style="margin-top: 0; margin-bottom: 0"><b><font color="#800000">&nbsp;</font><font color="#000080" size="2" face="Verdana">Total de Registros Encontrados :
                    </font> </b><font color="#000080" size="2" face="Verdana"><%=tem%></font></p>
                    <p style="margin-top: 0; margin-bottom: 0">
					<%
					end if
					%>
                    &nbsp;<font color="#000080" size="2" face="Verdana"><b>Total de
                    Apoiadores / Multiplicadores :
 </b> <%=igual%>
                    </font>
</form>
</body>

</html>

<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"

Server.ScriptTimeOut=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

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
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"
	
	set fonte = db.execute(ssql)
	
if fonte.eof=true then
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
                
  <table border="1" width="100%" bordercolor="#C0C0C0" cellspacing="0" cellpadding="2">
    <%
                  if fonte.eof=false then                  
                  ON ERROR RESUME NEXT
                  %>
    <tr> 
      <td width="15%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">ATRIBUIÇÃO</font></b></td>
      <td width="16%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">ÓRGÃO 
        APOIADO</font></b></td>
      <td width="20%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">NOME&nbsp;</font></b></td>
      <td width="14%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">LOTAÇÃO</font></b></td>
      <td width="10%" align="center" bgcolor="#C0C0C0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>ONDA</strong></font></td>
      <td width="11%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">MÓDULO</font></b></td>
      <td width="9%" align="center" bgcolor="#C0C0C0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>MOMENTO</strong></font></td>
      <td width="7%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">CHAVE</font></b></td>
      <td width="9%" align="center" bgcolor="#C0C0C0"><b><font face="Verdana" size="2">TELEFONE</font></b></td>
    </tr>
    <%
                  end if
                  %>
    <%
                  chave_ant=0
                  igual=0
                  
                  do until fonte.eof=true
                  
                  select case FONTE("ATRIBUICAO")
                  	CASE 1
                  		ATRIBUI="APOIADOR LOCAL"
                  	CASE 2
                  		ATRIBUI="MULTIPLICADOR"
                  END SELECT
                  
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
      <td width="15%" align="center"><font face="Verdana" size="1"><b><%=ATRIBUI%></b></font></td>
      <td width="16%" align="center"><font face="Verdana" size="1"><b><%=ORG_APOIO%></b></font></td>
      <td width="20%" align="center"><font face="Verdana" size="1"><b><%=UCASE(fonte("NOME"))%></b></font></td>
      <td width="14%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("LOTACAO"))%></font></td>
      <%
      set onda = db.execute("SELECT * FROM APOIO_LOCAL_ONDA WHERE USMA_CD_USUARIO='" & UCASE(fonte("CHAVE")) & "' AND APLO_NR_ATRIBUICAO=" & FONTE("ATRIBUICAO"))
      
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
      <td width="15%" align="center"><font face="Verdana" size="1"><%=sl_onda%></font></td>
      
      <td width="11%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("MODULO"))%></font></td>
	  <%
		MOMENTO=""
	
		if len(fonte("MOMENTO"))>1 AND fonte("MOMENTO")<>0 THEN
	  		MOMENTO = "1 e 2"
		else
	  		MOMENTO = fonte("MOMENTO")
		end if
		
		IF MOMENTO="" THEN
			MOMENTO=" - "
		END IF

	%>
      <td width="9%" align="center"><font face="Verdana" size="1"><%=MOMENTO%></font></td>
      <td width="7%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("CHAVE"))%></font></td>
      <td width="9%" align="center"><font face="Verdana" size="1"><%=UCASE(fonte("RAMAL"))%></font></td>
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

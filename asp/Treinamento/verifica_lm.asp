<%@LANGUAGE="VBSCRIPT"%>
<%
server.scripttimeout = 99999999
response.buffer = false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

if len(request("txtchave"))<>0 then
	ssql=""
	ssql="SELECT DISTINCT dbo.USUARIO_APROVADO.USAP_CD_USUARIO, dbo.USUARIO_APROVADO.CURS_CD_CURSO, "
	ssql=ssql+"dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO FROM dbo.USUARIO_APROVADO "
	ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO = dbo.USUARIO_APROVADO.USAP_CD_USUARIO "
	ssql=ssql+"INNER JOIN CURSO_FUNCAO ON "
	ssql=ssql+"dbo.USUARIO_APROVADO.CURS_CD_CURSO = dbo.CURSO_FUNCAO.CURS_CD_CURSO "
	ssql=ssql+"WHERE dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO<>'AP'"
	ssql=ssql+"AND dbo.USUARIO_APROVADO.USAP_CD_USUARIO='" & trim(ucase(REQUEST("txtchave"))) & "'"	
	ssql=ssql+"ORDER BY dbo.USUARIO_APROVADO.USAP_CD_USUARIO, dbo.USUARIO_APROVADO.CURS_CD_CURSO,dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO"
else
	ssql=""
	ssql="SELECT DISTINCT dbo.USUARIO_APROVADO.USAP_CD_USUARIO, dbo.USUARIO_APROVADO.CURS_CD_CURSO, "
	ssql=ssql+"dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO FROM dbo.USUARIO_APROVADO "
	ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO = dbo.USUARIO_APROVADO.USAP_CD_USUARIO "
	ssql=ssql+"INNER JOIN CURSO_FUNCAO ON "
	ssql=ssql+"dbo.USUARIO_APROVADO.CURS_CD_CURSO = dbo.CURSO_FUNCAO.CURS_CD_CURSO "
	ssql=ssql+"WHERE dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO<>'AP' "

	str01 = request("str01")
	str02 = request("str02")
	str03 = request("str03")
	
	orgao=""
	
	if str01<>0 then
		orgao = str01
	end if

	if str02<>"000" then
		orgao = str02
	end if

	if str03<>0 then
		orgao = str03
	end if
	
	if len(orgao)>0 then
		ssql = ssql + "AND dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' " 
	end if
		
	funcneg = trim(request("txtfuncao"))
	
	if funcneg="'0'," then
		funcneg=""
	end if
	
	if len(funcneg)>1 then
		funcneg = left(funcneg,len(funcneg)-1)
	end if
	
	if len(funcneg)>0 then
		ssql = ssql + "AND dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO IN (" & funcneg & ")" 
	end if

	ssql=ssql+"ORDER BY dbo.USUARIO_APROVADO.USAP_CD_USUARIO, dbo.USUARIO_APROVADO.CURS_CD_CURSO,dbo.USUARIO_APROVADO.USAP_TX_APROVEITAMENTO"
end if

set fonte = db.execute(ssql)

reg = fonte.RecordCount
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">
<form name="frm1" method="POST" action="valida_lm.asp">

  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="border-collapse: collapse" bordercolor="#111111">
    <tr> 
      <td width="158" height="20" colspan="2">&nbsp;</td>
      <td width="577" height="60" colspan="3">&nbsp;</td>
      <td width="150" valign="top" colspan="2"> <table width="150" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
            <td bgcolor="#330099" width="51" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="49" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="50" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr>
            <td bgcolor="#330099" height="12" width="51" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="49" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="50" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="38">&nbsp; </td>
      <td height="20" width="120"> <p align="right">
      <%if reg>0 then%>
      <img border="0" src="../../imagens/confirma_f02.gif" onclick="document.frm1.submit()"> 
      <%end if%>
      </td>
      <%if reg>0 then%>
      <td height="20" width="181"> <font size="2" face="Verdana" color="#000080"><b>&nbsp;Enviar</b></font>
      <%end if%> 
      </td>
      <td height="20" width="44">&nbsp;</td>
      <td height="20" width="352">&nbsp;</td>
      <td height="20" width="84">&nbsp; </td>
      <td height="20" width="66">&nbsp; </td>
    </tr>
  </table>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">

<input type="hidden" name="txtQuery" size="69" value="<%=ssql%>"></p>

  <table border="0" width="88%" height="56">
    <tr>
      <td width="6%" height="29">&nbsp;</td>
      <td width="10%" height="29">&nbsp;</td>
      <td width="42%" height="29"><font face="Verdana" color="#000080">Treinamento - Liberação Manual de Usuários em Cursos (<b>LM</b>)</font></td>
    </tr>
    <tr>
      <td width="6%" height="26"></td>
      <td width="10%" height="26">
      	<input type="hidden" name="txtmotivo" size="6" value="<%=request("selMotivo")%>">
      </td>
      <td width="42%" height="26"><p align="left">
      <%if reg>0 then%>
      <font face="Verdana" color="#000080" size="2">Total <b><%=reg%></b> Registro(s)</font>
      <%end if%>
      </td>
    </tr>
    </table>
  <p align="left">
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="79%" id="AutoNumber1" height="51">
             <%
             if reg > 0 then
             %>
             <tr>
             <td width="22%" height="23">&nbsp;</td>
             <td width="25%" height="23" bgcolor="#000080"><b><font size="2" face="Verdana" color="#EFEFEF">Usuário</font></b></td>
             <td width="29%" height="23" bgcolor="#000080"><b><font size="2" face="Verdana" color="#EFEFEF">Curso</font></b></td>
             <td width="24%" height="23" bgcolor="#000080"><p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" face="Verdana" color="#EFEFEF">Liberado Manualmente</font></b><p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2" face="Verdana" color="#EFEFEF">
             (LM)</font></b></td>
             </tr>
             <%
             end if
             i=0
             ch_ant = ""
             do until i = reg
             
             ch_atual = fonte("USAP_CD_USUARIO")
             if ch_ant = ch_atual then
             	chave=""
             else
	            chave = fonte("USAP_CD_USUARIO")
	            if cor="white" then
	            	cor="#E2E2E2"
	            else
	            	cor="white"
	            end if
			 end if             
             %>
             <tr>
                        <td width="22%" height="27">&nbsp;</td>
                        <td width="25%" height="27" bgcolor="<%=cor%>"><font size="2" face="Verdana"><b><%=chave%></b></font></td>
                        <td width="29%" height="27" bgcolor="<%=cor%>"><font size="2" face="Verdana"><%=fonte("CURS_CD_CURSO")%></font></td>
                        <%
							if fonte("USAP_TX_APROVEITAMENTO") = "LM" then
								checado = "checked"
							else
								checado = ""						
							end if
                        %>
                        <td width="24%" height="27" bgcolor="<%=cor%>"><p align="center"><input type="checkbox" name="<%=fonte("USAP_CD_USUARIO")%>_<%=fonte("CURS_CD_CURSO")%>" value="1" <%=checado%>></td>
             </tr>
			 <%
			 i = i + 1
			 ch_ant = fonte("USAP_CD_USUARIO")
			 fonte.movenext
			 loop
			 %>
  </table>
</form>
<%
db.close
set db = nothing
if i = 0 then
%> 
<p align="center"><b><font color="#800000">Nenhum Registro Encontrado para a Seleção!</font></b></p>
<%
end if
%>
</body>
</html>
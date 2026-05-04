<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
COMPL=" AND MEPR_CD_MEGA_PROCESSO=" & mega

str_Assunto=0
str_Assunto=request("selAssunto")

if str_Assunto<>0 then
		compl="  AND SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
end if

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE ONDA_CD_ONDA<>4" & COMPL &" ORDER BY CENA_CD_CENARIO"

set rs=db.execute(ssql)
%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="../gera_rel_megaemp.asp">
              <input type="hidden" name="txtEmpSelecionada"><input type="hidden" name="txtOpc" value="<%=str_Opc%>">
<%if request("excel")<>1 then%>              
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr>
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <%end if%>
              <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
              
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="3" color="#330099">Relatório 
    de Problemas com status de Cenários</font></p>
              <%
              SET RSMEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)
              %>
              <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#330099" size="2">Mega-Processo
              - <%=mega%> - <%=rsmega("MEPR_TX_dESC_MEGA_PROCESSO")%></font></b></p>
              <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
              <table border="0" width="95%" cellpadding="3" cellspacing="1">
                <tr>
                  <td width="51%" bgcolor="#330099">
                    <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#FFFFFF">Cenário</font></b></td>
                  <td width="17%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Status</font></b></td>
                  <td width="32%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Problema
                    Ocorrido</font></b></td>
                </tr>
                <%
                tem=0
                do until rs.eof=true
                PROBLEMA=""
                exibe=0
                VALOR=rs("CENA_CD_CENARIO")
                
                VAL_STATUS=RS("CENA_TX_SITUACAO")
                
                SELECT CASE VAL_STATUS
                	
                CASE "DS"
                		IF ISNULL(rs("CENA_TX_SITU_DESENHO_TIPO")) THEN
                			PROBLEMA="Cenário Desenhado sem Informação sobre Desenvolvimento"
                			Exibe=1
                end if
                
                CASE "TD"
                		ONDA=LEFT(RIGHT(VALOR,8),3)	
                		IF ONDA<>"PMA" AND ONDA<>"PSA" THEN
                			PROBLEMA="Cenário TESTADO NO PED não pertence às ondas PMA e PSA"
                			Exibe=1
                end if
                
                CASE "TQ"
		       		ONDA=LEFT(RIGHT(VALOR,8),3)	
                		IF ONDA<>"PMA" AND ONDA<>"PSA" THEN
                			PROBLEMA="Cenário TESTADO NO PEQ não pertence às ondas PMA e PSA"
                			Exibe=1
                end if
                
                CASE "PT"
                		IF ISNULL(rs("CENA_TX_SITU_DESENHO_TIPO")) THEN
                			exibe=1
                			PROBLEMA="Cenário PRONTO PARA TESTE sem Informação sobre Desenvolvimento"
                		ELSE
                			IF rs("CENA_TX_SITU_DESENHO_TIPO")=1 AND ISNULL(rs("CENA_TX_SITU_DESENHO_DESE")) OR ISNULL(rs("CENA_TX_SITU_DESENHO_CONF")) THEN
                				exibe=1
			         			PROBLEMA="Cenário PRONTO PARA TESTE com informação sobre Desenvolvimento incompleta"
			         		ELSE
			         			IF rs("CENA_TX_SITU_DESENHO_TIPO")=2 AND ISNULL(rs("CENA_TX_SITU_DESENHO_CONF")) THEN
			         				exibe=1
				         			PROBLEMA="Cenário PRONTO PARA TESTE com informação sobre Desenvolvimento incompleta"
				         		END IF
							END IF
			         	END IF
                
                CASE "TQ"
                		
                		IF LEN(PROBLEMA)>0 THEN
                			PROBLEMA=PROBLEMA & " - " 
                		END IF
                		
                		IF ISNULL(rs("CENA_TX_SITU_DESENHO_TIPO")) THEN
                			exibe=1
                			PROBLEMA = PROBLEMA & "Cenário TESTADO NO PEQ sem Informação sobre Desenvolvimento"
                		ELSE
                			IF rs("CENA_TX_SITU_DESENHO_TIPO")=1 AND ISNULL(rs("CENA_TX_SITU_DESENHO_DESE")) OR ISNULL(rs("CENA_TX_SITU_DESENHO_CONF")) THEN
                				exibe=1
			         			PROBLEMA="Cenário TESTADO NO PEQ com informação sobre Desenvolvimento incompleta"
			         		ELSE
			         			IF rs("CENA_TX_SITU_DESENHO_TIPO")=2 AND ISNULL(rs("CENA_TX_SITU_DESENHO_CONF")) THEN
			         				exibe=1
				         			PROBLEMA=PROBLEMA & "Cenário TESTADO NO PEQ com informação sobre Desenvolvimento incompleta"
				         		END IF
							END IF
			         	END IF
                
                end select
                
                IF EXIBE=1 THEN

                IF COR="WHITE" THEN
                	COR="#DDDDDD"
                ELSE
                	COR="WHITE"
                END IF
                
                %>
                <tr>
                  <td width="51%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=VALOR & "-" & RS("CENA_TX_TITULO_CENARIO")%></font></td>
                  <%tem=1%>
                  <%
                  SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "SITUACAO_GERAL WHERE SITU_TX_CD_STATUS='" & RS("CENA_TX_SITUACAO") & "'")
                  VALOR_STATUS=TEMP("SITU_TX_DESC_SITUACAO")                  
                  %>
                  <td width="17%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=VALOR_STATUS%></font></td>
                  <td width="32%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=PROBLEMA%></font></td>
                </tr>
				  <%
				  END IF
				  rs.movenext
				  loop
				  %>	                
              </table>
              <p style="margin-top: 0; margin-bottom: 0">&nbsp;
              <%if tem=0 then%>
              <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#800000"><b>Nenhum
              Registro Encontrado para a Seleção</b></font></p>
              <%end if%>
              
</form>
</body>
</html>

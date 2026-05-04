<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

macro=request("macro")
mega=request("mega")

if len(macro)<>0 then
	ssql="SELECT * FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & macro  & " order by MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO"
else
	if len(mega)<>0 then
		ssql="SELECT * FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA WHERE MEPR_CD_MEGA_PROCESSO=" & mega  & " order by MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO"
	end if
end if

if len(macro)<>0 and len(mega)<>0 then
	ssql="SELECT * FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & macro  & " and MEPR_CD_MEGA_PROCESSO=" & mega  & " order by MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO"
end if

'response.write ssql

set rs=conn_db.execute(ssql)

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Histórico de Validação</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#990000" vlink="#990000" alink="#990000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="http://its_server3/valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1">
  <table width="522" border="0">
    <tr>
      <td width="338"><font color="#330099" face="Verdana" size="3">Visualização de
        Andamento de Validação</font></td>
      <td width="170"><div align="right"><strong><font color="#990000" size="3"><a href="javascript:window.close()">Fechar 
          Janela</a></font></strong></div></td>
    </tr>
  </table>
		<%
       SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & macro )
       %>
        <p align="left"><font color="#330099" face="Verdana" size="2"><b>Macro-Perfil
        Selecionado :&nbsp; <%=macro%> </b>- <%=TEMP("MCPE_TX_NOME_TECNICO")%> </font></p>
        
  <table border="0" width="573">
    <tr> 
      <td width="233" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Mega-Processo 
        a Aprovar</b></font></td>
      <td width="160" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Transação</b></font></td>
      <td width="166" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Status
        Atual</b></font></td>
    </tr>
    <%
          tem=0
          do until rs.eof=true
		  
	       valor_mega=""
          
          select case rs("MAOA_TX_AUTORIZADO")
          			
			case "0"
				valor="Ainda não validado"
			case "1"
				valor="Aprovado"
			case "2"
				valor="Reprovado"
			case "3"
				valor="Aprovação Diferenciada"
			end select
			
			SET MEGA_ = CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO"))
			IF MEGA_.EOF=FALSE THEN
				VALOR_MEGA=MEGA_("MEPR_TX_DESC_MEGA_PROCESSO")
			END IF
			
			IF COR="WHITE" THEN
				COR="#DDDDDD"
			ELSE
				COR="WHITE"
			END IF
          %>
    <tr> 
      <td width="233" align="center" bgcolor="<%=COR%>"> <font face="Verdana" size="1"><%=valor_mega%></font></td>
      <td width="160" align="center" bgcolor="<%=COR%>"> <font face="Verdana" size="1"><%=rs("TRAN_CD_TRANSACAO")%></font></td>
      <td width="166" align="center" bgcolor="<%=COR%>"> <font face="Verdana" size="1"><%=valor%></font></td>
    </tr>
    <%
          tem = tem + 1
          rs.movenext
          loop
          %>
  </table>
        
  <div align="left">
<%if tem=0 then%>
    <font color="#800000"><b> Nenhum Registro Encontrado!</b></font> 
    <%end if%>
    <input type="hidden" name="txtcaminho" size="20">
  </div>
</form>
</body>
</html>

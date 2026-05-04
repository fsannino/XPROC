<%@LANGUAGE="VBSCRIPT"%>

<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

macro=request("macro")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & macro  & " order by ATUA_DT_ATUALIZACAO" )

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
  <table width="84%" border="0">
    <tr>
      <td><font color="#330099" face="Verdana" size="3">Visualização de Histórico</font></td>
      <td><div align="right"><strong><font color="#990000" size="3"><a href="javascript:window.close()">Fechar 
          Janela</a></font></strong></div></td>
    </tr>
  </table>
  <%
        SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & macro )
        %>
        <p align="left"><font color="#330099" face="Verdana" size="2"><b>Macro-Perfil
        Selecionado :&nbsp; <%=macro%> </b>- <%=TEMP("MCPE_TX_NOME_TECNICO")%> </font></p>
        
  <table border="0" width="85%" height="35">
    <tr> 
      <td width="14%" bgcolor="#330099" align="center" height="18"><font color="#FFFFFF" size="2" face="Verdana"><b>Usuário</b></font></td>
      <td width="27%" bgcolor="#330099" align="center" height="18"><font color="#FFFFFF" size="2" face="Verdana"><b>Data 
        / Hora</b></font></td>
      <td width="24%" bgcolor="#330099" align="center" height="18"><font color="#FFFFFF" size="2" face="Verdana"><b>Status</b></font></td>
      <td width="25%" bgcolor="#330099" align="center" height="18"><font color="#FFFFFF" size="2" face="Verdana"><b>Mega-Processo</b></font></td>
    </tr>
    <%
     tem=0
     
     do until rs.eof=true
		  
			valor_mega=""
		  
	  		ssql=""
			ssql="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO")
			
			SET MEGA_ = CONN_DB.EXECUTE(ssql)
			
			IF MEGA_.EOF=FALSE THEN
				VALOR_MEGA=MEGA_("MEPR_TX_DESC_MEGA_PROCESSO")
			END IF
             
          select case rs("MHVA_TX_SITUACAO_MACRO")

			case "EE"
				VALOR="Em Elaboração"
			case "AT"
				valor="Transação Alterada"
			case "EA"
				VALOR="Em Aprovação"
			case "NA"
				VALOR="Não Aprovado"
			case "EC"
				VALOR="Em Criação no R/3"
			case "RE"
				valor="Recusado no R/3"
			case "EX"
				valor="Função Excluída"
			case "MR"
				valor="Mudado para Referência"          			
			case "EL"
				valor="Excluído"				
			case "CR"			
				VALOR="Criado no R/3"
			case "AR"
				VALOR="Em alteração no R/3"				
			case "ER"
				VALOR="Em exclusão no R/3"				
			case "AP"
				VALOR="Alterado no R/3"				
			case "EP"
				VALOR="Excluido no R/3"							
			case "RD"
				valor="Aprovação Diferenciada"
			end select
	       %>
    <tr> 
      <td width="14%" align="center" height="15"> <font size="1" face="Verdana"> <%=rs("ATUA_CD_NR_USUARIO")%></font></td>
      <td width="27%" align="center" height="15"> <font size="1" face="Verdana"> <%=FORMATDATETIME((rs("ATUA_DT_ATUALIZACAO")),0)%></font></td>
      <td width="24%" align="center" height="15"> <font size="1" face="Verdana"> <%=VALOR%></font></td>
      <td width="25%" align="center" height="15"> <font size="1" face="Verdana"> <%=VALOR_MEGA%></font></td>
    </tr>
    <tr bgcolor="#CCCCCC"> 
      <td height="1" align="center"><strong><font color="#000066" size="1" face="Verdana">Coment&aacute;rios 
        : </font></strong></td>
      <td colspan="3" align="center" height="1"> 
        <div align="justify"><font size="1" face="Verdana"><%=rs("MHVA_TX_COMENTARIO")%></font></div></td>
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

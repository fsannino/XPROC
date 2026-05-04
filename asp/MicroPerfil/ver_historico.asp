<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

micro=request("micro")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro  & "' order by ATUA_DT_ATUALIZACAO")

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

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="http://its_server3/valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
        <p align="left"><font color="#330099" face="Verdana" size="3">Visualização
        de Histórico</font></p>
        <%
        SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")
        %>
        <p align="left"><font color="#330099" face="Verdana" size="2"><b>Micro-Perfil
        Selecionado :&nbsp; <%=micro%> </b>- <%=TEMP("MICR_TX_DESC_MICRO_PERFIL")%> </font></p>
        <table border="0" width="87%">
          <tr>
            <td width="1%" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Usuário</b></font></td>
            <td width="26%" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Data
              / Hora</b></font></td>
            <td width="25%" bgcolor="#330099" align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Status</b></font></td>
          </tr>
          <%
          tem=0
          do until rs.eof=true
          
          select case rs("MHVA_TX_SITUACAO_MICRO")
          			
			case "EE"
				VALOR="Em Elaboração"
			case "EL"
				VALOR="Excluído"
			case "RE"
				VALOR="Recusado no R/3"
				
			case "EC"
				VALOR="Em Criação no R/3"
			case "CR"
				VALOR="Criado no R/3"
				
			end select

          %>
          <tr>
            <td width="1%" align="center">
              <font size="1" face="Verdana">
              <%=rs("ATUA_CD_NR_USUARIO")%></font></td>
            <td width="26%" align="center">
              <font size="1" face="Verdana">
              <%=FORMATDATETIME((rs("ATUA_DT_ATUALIZACAO")),0)%></font></td>
            <td width="25%" align="center">
              <font size="1" face="Verdana">
              <%=VALOR%></font></td>
          </tr>
          <tr>
            <td width="1%" align="center" bgcolor="#CCCCCC">
              <b><font face="Verdana" size="1">Comentários :</font></b></td>
            <td width="51%" align="center" colspan="2" bgcolor="#CCCCCC">
              <p align="left"><b><font face="Verdana" size="1"><%=rs("MHVA_TX_COMENTARIO")%></font></b></td>
          </tr>
          <%
          tem = tem + 1
          rs.movenext
          loop
          %>
          
         </table>
        <%if tem=0 then%>
        <font color="#800000"><b>
        Nenhum Registro Encontrado!</b></font>
        <%end if%>
        <input type="hidden" name="txtcaminho" size="20">
  </form>
</body>
</html>

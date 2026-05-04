<%@LANGUAGE="VBSCRIPT"%> 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO=" & request("Mega")  & " AND (MICR_TX_SITUACAO = 'EE' OR  MICR_TX_SITUACAO = 'RE' OR  MICR_TX_SITUACAO = 'EC')")
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<SCRIPT LANGUAGE="JavaScript">

function Confirma()
{
document.frm1.submit();
}

function addbookmark()
{
bookmarkurl="http://www.sinergia.petrobras.com.br/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_micro1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60"><font size="1" color="#FFFFFF"><b><%=Conn_String_Cogest_Gravacao%></b></font></td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://http://www.sinergia.petrobras.com.br/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">&nbsp;
      
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Elaboração -&gt; Em Criação</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Status
        Alterado com Sucesso</font></b></p>
        <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td height="41" width="176" align="right"><a href="selec_valida_micro1.asp"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela de E</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">ncaminhamento
              de Status</font></td>
          </tr>
          <tr>
            <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
          </tr>
        </table>
        <%
        	do until rs.eof=true
        
          	VALOR=REQUEST("micro_" & trim(rs("MICR_TX_SEQ_MICRO_PERFIL")))
          	
          	COMENTARIO=REQUEST("coment_" & trim(rs("MICR_TX_SEQ_MICRO_PERFIL")))
          	
          SELECT CASE VALOR
        	
        	CASE 1
        		VALOR_STATUS="EE"
        	CASE 2
        		VALOR_STATUS="NA"
        	CASE 3
        		VALOR_STATUS="EA"
			case 4
				VALOR_STATUS="AT"
			case 6
				VALOR_STATUS="EC"
			END SELECT
			
        	SSQL=""
        	SSQL="SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL WHERE MICR_TX_SEQ_MICRO_PERFIL='" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "'"
        	
      		SET ATUAL = CONN_DB.EXECUTE(SSQL)
      		
      		VALOR_ATUAL = ""
      		VALOR_ATUAL = ATUAL("MICR_TX_SITUACAO")
      		
      		
			IF TRIM(VALOR_ATUAL)<>TRIM(VALOR_STATUS) THEN
      
	  			CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MICRO_PERFIL SET MICR_TX_SITUACAO='" &  VALOR_STATUS & "', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MICR_TX_SEQ_MICRO_PERFIL='" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "'")
	
				SSQL=""
				SSQL="SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL='" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "'"
				
				SET HIST = CONN_DB.EXECUTE(SSQL)
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        			if atual > 1 then
	        			atual = atual
   		     		else
       	 			atual=1
        			end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MICR_TX_SEQ_MICRO_PERFIL, MHVA_TX_SITUACAO_MICRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & ATUAL & ", '" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "', '"& VALOR_STATUS &"', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
        		conn_db.execute(ssql)
        		
        		set valida=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "GRUPO_VALIDADOR WHERE MEPR_CD_MEGA_PROCESSO=" & request("MEGA") & " AND GRVA_TX_ASSUNTO='MIC' AND GRVA_TX_TIPO2='G1'")
				
				do until valida.eof=true
				
					set correio = server.CreateObject("Persits.MailSender")

					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"	     			
	     			
					valor = valida("USUA_CD_USUARIO") & "@petrobras.com.br"
	     			
	     			correio.AddAddress valor
        			correio.Subject="Alteração de Status de Micro-Perfil"
        			
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Micro-Perfil '" & rs("MICR_TX_DESC_MICRO_PERFIL") & "' foi alterado para EM APROVAÇÃO - EM : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					correio.send
				
					valida.movenext
				
				loop

				
        	END IF
        	
        rs.movenext
        
        loop
        %>
  </form>
</body>
</html>
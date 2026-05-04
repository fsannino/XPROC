<%@LANGUAGE="VBSCRIPT"%> 
<%
server.scripttimeout=99999999

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ") AND (MCPE_TX_SITUACAO='EC' OR MCPE_TX_SITUACAO = 'ER' OR MCPE_TX_SITUACAO = 'AR') ORDER BY MEPR_CD_MEGA_PROCESSO, MCPR_NR_SEQ_MACRO_PERFIL")
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
<form name="frm1" method="post" action="valida_status5.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://http://www.sinergia.petrobras.com.br/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
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
        de Status :&nbsp; Em Criação -&gt; Criado no R/3</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Status
        Alterados com Sucesso</font></b></p>
        <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td height="41" width="176" align="right"><a href="verifica_valida_status5.asp"><img src="selecao_F02.gif" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela de Encaminhamento de Status</font></td>
          </tr>
          <tr>
            <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="selecao_F02.gif" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
          </tr>
        </table>
        <%
		COMENTARIO=""
		
		VALOR_ATUAL=""
		VALOR_STATUS=""
		
		do until rs.eof=true
		
		VALOR_ATUAL=RS("MCPE_TX_SITUACAO")
		
		VALOR = REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))
		COMENTARIO=REQUEST("coment_" & trim(rs("MCPE_TX_NOME_TECNICO")))
		MEGA_=REQUEST("MEGA_" & trim(rs("MCPE_TX_NOME_TECNICO")))
		
		VALOR_STATUS=""
		
		SELECT CASE VALOR
			CASE 10
			    ' ALTERADO NO R3
				VALOR_STATUS="AP"	
			CASE 11
			    ' EXCLUIDO NO R3 
				VALOR_STATUS="EP"
			CASE 1
			    ' CRIADO NO R3 
				VALOR_STATUS="CR"
			CASE 2
			    ' RECUSADO 
				VALOR_STATUS="RE"
		END SELECT
		
		IF VALOR<>"" THEN
		
		IF VALOR_ATUAL<>VALOR_STATUS THEN
	           if VALOR_STATUS= "CR" OR VALOR_STATUS = "AP" then		
               str_SQL = ""
				   str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
				   str_SQL = str_SQL & " SET MCPT_NR_SITUACAO_PROCESSAMENTO = 1 "
				   str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
				   str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO = GETDATE() "
					   str_SQL = str_SQL & " , ATUA_TX_OPERACAO = 'A' "
				   str_SQL = str_SQL & "  WHERE MCPR_NR_SEQ_MACRO_PERFIL='" & rs("MCPR_NR_SEQ_MACRO_PERFIL") & "'"
			   CONN_DB.EXECUTE(str_SQL)
			end if
			
			if REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))=10 then
			
				val_verifica = "ALTERADO NO R/3"
        
				CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='AP', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        		if atual > 1 then
        			atual = atual
        		else
        			atual=1
        		end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & rs("MEPR_CD_MEGA_PROCESSO") & ",'" & COMENTARIO & "', " & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'AP', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
        		'RESPONSE.WRITE SSQL
        		
        		conn_db.execute(ssql)
        		       		
	        		set correio = server.CreateObject("Persits.MailSender")

					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
	     			set destino = conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
     			
					valor1 = destino("MEPR_TX_VALIDA_MACRO") & "@petrobras.com.br"
					valor2 = rs("ATUA_CD_NR_USUARIO") & "@petrobras.com.br"
						     			
	     			correio.AddAddress valor1
	     			if trim(valor2)<>"" then			
						correio.AddAddress valor2
					end if
	     			
        			correio.Subject="Alteração de Status de Macro-Perfil"
					
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para ALTERADO NO R/3 : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					ON ERROR RESUME NEXT
					correio.send
					ERR.CLEAR

        		SET CORREIO=NOTHING

        	END IF

			IF REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))=11 then
			
				val_verifica = "EXCLUÍDO NO R/3"
        
				CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='EP', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        		if atual > 1 then
        			atual = atual
        		else
        			atual=1
        		end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & rs("MEPR_CD_MEGA_PROCESSO") & ",'" & COMENTARIO & "', " & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'EP', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
        		'RESPONSE.WRITE SSQL
        		
        		conn_db.execute(ssql)
        		       		
	        		set correio = server.CreateObject("Persits.MailSender")
					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
	     			set destino = conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
	     			
					valor1 = destino("MEPR_TX_VALIDA_MACRO") & "@petrobras.com.br"
					valor2 = rs("ATUA_CD_NR_USUARIO") & "@petrobras.com.br"
	     			
	     			correio.AddAddress valor1
	     			if trim(valor2)<>"" then			
						correio.AddAddress valor2
					end if
	     			
        			correio.Subject="Alteração de Status de Macro-Perfil"
					
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para EXCLUIDO NO R/3 : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					ON ERROR RESUME NEXT
					correio.send
					ERR.CLEAR

        		SET CORREIO=NOTHING

        	END IF

        	IF REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))=1 then
        
				val_verifica = "CRIADO NO R/3"
				
				CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='CR', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
				'CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='CR', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MCPE_TX_NOME_TECNICO='" & rs("MCPE_TX_NOME_TECNICO") & "'")
        		
        		SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        		if atual > 1 then
        			atual = atual
        		else
        			atual=1
        		end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & rs("MEPR_CD_MEGA_PROCESSO") & ",'" & COMENTARIO & "', " & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'CR', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
        		'RESPONSE.WRITE SSQL
        		
        		conn_db.execute(ssql)
        		       		
	        		set correio = server.CreateObject("Persits.MailSender")
					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
	     			set destino = conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
	     			
					valor1 = destino("MEPR_TX_VALIDA_MACRO") & "@petrobras.com.br"
					valor2 = rs("ATUA_CD_NR_USUARIO") & "@petrobras.com.br"
	     			
	     			correio.AddAddress valor1
	     			if trim(valor2)<>"" then			
						correio.AddAddress valor2
					end if
	     			
        			correio.Subject="Alteração de Status de Macro-Perfil"
					
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para CRIADO NO R/3 : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					ON ERROR RESUME NEXT
					correio.send
					ERR.CLEAR

        		SET CORREIO=NOTHING

        	END IF
			
			IF REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))=2 then
			
				val_verifica = "RECUSADO NO R/3"
        	
        		CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='RE', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_DT_ATUALIZACAO=GETDATE() WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        		if atual > 1 then
        			atual = atual
        		else
        			atual=1
        		end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & rs("MEPR_CD_MEGA_PROCESSO") & ",'" & COMENTARIO & "', " & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'RE', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
        		'RESPONSE.WRITE SSQL
        		
        		conn_db.execute(ssql)
        		       		
	        		set correio = server.CreateObject("Persits.MailSender")
					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
	     			set destino = conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
	     			
					valor1 = destino("MEPR_TX_VALIDA_MACRO") & "@petrobras.com.br"
					valor2 = rs("ATUA_CD_NR_USUARIO") & "@petrobras.com.br"
	     			
	     			correio.AddAddress valor1
	     			if trim(valor2)<>"" then			
						correio.AddAddress valor2
					end if
	     			
        			correio.Subject="Alteração de Status de Macro-Perfil"
					
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para CRIADO NO R/3 : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					ON ERROR RESUME NEXT
					correio.send
					ERR.CLEAR
					
					SSQL1="SELECT * FROM " & Session("PREFIXO") & "GRUPO_VALIDADOR WHERE MEPR_CD_MEGA_PROCESSO=" & MEGA_ & " AND GRVA_TX_ASSUNTO='MAC' AND GRVA_TX_TIPO2='G1'"
					set valida=conn_db.execute(SSQL1)
				
				do until valida.eof=true
				
					set correio = server.CreateObject("Persits.MailSender")
					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
					valor = valida("USUA_CD_USUARIO") & "@petrobras.com.br"
						     			
					if trim(valor)<>"" then			
						correio.AddAddress valor
					end if
					
        			correio.Subject="Alteração de Status de Macro-Perfil"
        			
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para " & VAL_VERIFICA & " : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					ON ERROR RESUME NEXT
					correio.send
					ERR.CLEAR
				
					valida.movenext
				
				loop					

        		SET CORREIO=NOTHING
        		
        	END IF
    	
        	END IF
        	
        	END IF
			
        rs.movenext
        
        loop
    %>
  </form>
</body>
</html>
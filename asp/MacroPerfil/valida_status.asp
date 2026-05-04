<%@LANGUAGE="VBSCRIPT"%> 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

opto=request("opt")
str_origem=request("txtOrigem")
str_Acao = request("Acao")
IF opto="EC" then
	req1="EM CRIAÇÃO NO R/3"
	req2="G3"
else
	req1="EM APROVAÇÃO"
	req2="G1"
end if

set rs=conn_db.execute("SELECT * FROM MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & request("MACRO"))
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
        de Status</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Status
        do Macro-Perfil Alterado com Sucesso</font></b></p>
        
  <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="selecao_F02.gif" border="0"></a></td>
      <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela Principal</font></td>
    </tr>
    <% if str_Origem = 0 then 
	      if str_Acao = "C" then
	%>	   
    <tr> 
      <td height="41" align="right"><a href="incluir_macro_perfil.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela de cria&ccedil;&atilde;o de Macro Perfil</font></td>
    </tr>
	<% end if 
	   if str_Acao = "M" then %>
    <tr> 
      <td height="41" align="right"><a href="seleciona_macro_perfil.asp?pOPT=1"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela de altera&ccedil;&atilde;o de Macro Perfil</font></td>
    </tr>	
    <% end if
	end if %>
    <tr> 
      <td height="41" align="right">&nbsp;</td>
      <td height="41">&nbsp;</td>
    </tr>
  </table>
        <%
      if request("acao")="C" then
			COMENTARIO="ENCAMINHAMENTO VIA INCLUSÃO"
	 	else
		 	COMENTARIO="ENCAMINHAMENTO VIA MODIFICAÇÃO"
		 end if
	
		CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='" & opto & "' WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
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
        		SSQL=SSQL+"VALUES(" & rs("MEPR_CD_MEGA_PROCESSO") & ",'" & COMENTARIO & "', " & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", '" & opto & "', 'I', '" & Session("CdUsuario") & "', GETDATE())"
        		
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
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para " & REQ1 & " : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
					correio.send
					
					SSQL1="SELECT * FROM " & Session("PREFIXO") & "GRUPO_VALIDADOR WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO") & " AND GRVA_TX_ASSUNTO='MAC' AND GRVA_TX_TIPO2='" & REQ2 & "'"
									
					set valida=conn_db.execute(SSQL1)
					
					'RESPONSE.WRITE VALIDA.EOF
					
				
				do until valida.eof=true
				
					set correio = server.CreateObject("Persits.MailSender")

					correio.host = "164.85.62.165"
	     			correio.from="xproc@S600146.petrobras.com.br"
	     			correio.fromname="Suporte XPROC"
	     				     			
					valor = valida("USUA_CD_USUARIO") & "@petrobras.com.br"	 	     			

	     			correio.AddAddress valor
  	    			correio.Subject="Alteração de Status de Macro-Perfil"
        			
					data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
        			correio.Body=" O Status do Macro-Perfil '" & rs("MCPE_TX_NOME_TECNICO") & "' foi alterado para " & REQ1 & " : " & DATA_ATUAL & " / POR : " & Session("CdUsuario")
					
					correio.send
				
					valida.movenext

				loop					

        		SET CORREIO=NOTHING
        	%>
  </form>
</body>
</html>
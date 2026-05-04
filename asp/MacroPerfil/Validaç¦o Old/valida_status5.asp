<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO=" & request("mega") & " AND MCPE_TX_SITUACAO='EC'")
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
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
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
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Cenario/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Cenario/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Cenario/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Cenario/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Cenario/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Cenario/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      &nbsp;
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Criação -&gt; Criado no R/3</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Status
        Alterados com Sucesso</font></b></p>
        <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td height="41" width="176" align="right"><a href="selec_valida_status5.asp"><img src="file:///G:/xproc/imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela de Encaminhamento de Status</font></td>
          </tr>
          <tr>
            <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="file:///G:/xproc/imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
          </tr>
        </table>
        <%do until rs.eof=true
        	IF REQUEST("macro_" & trim(rs("MCPE_TX_NOME_TECNICO")))=1 then
        		CONN_DB.EXECUTE("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='CR' WHERE MCPE_TX_NOME_TECNICO='" & rs("MCPE_TX_NOME_TECNICO") & "'")
        		
        		SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
        		
        		ATUAL=HIST("CODIGO")
        		ATUAL = ATUAL + 1
        		
        		if atual > 1 then
        			atual = atual
        		else
        			atual=1
        		end if
        		
        		SSQL=""
        		SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
        		SSQL=SSQL+"VALUES(" & ATUAL &", " & rs("MCPR_NR_SEQ_MACRO_PERFIL") & ", 'CR', 'I', '" & session("cdUsuario") & "', GETDATE())"
        		
        		conn_db.execute(ssql)

        	END IF
        rs.movenext
        loop
        %>
  </form>
</body>
</html>

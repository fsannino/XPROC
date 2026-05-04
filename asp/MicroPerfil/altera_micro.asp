<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

micro = request("selMicroPerfil")

set rs=db.execute("SELECT * FROM MICRO_PERFIL WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function pega_tamanho()
{
document.frm1.txttamanho.value=document.frm1.txtDescM.value.length
if (document.frm1.txtDescM.value.length > 61) {
	str1=document.frm1.txtDescM.value;
	document.frm1.txtDescM.value=str1.slice(0,61);
	document.frm1.txttamanho.value=str2.length;
}
}
</script>

<script>
function Confirma()
{
   if(document.frm1.txtDescM.value == "")
      {
      alert("É obrigatória o preenchimento do campo Descrição.");
      document.frm1.txtDescM.focus();
      return;
      }
   if(document.frm1.txtdetalM.value == "")
      {
      alert("É obrigatória o preenchimento do campo Descrição Detalhada.");
      document.frm1.txtdetalM.focus();
      return;
      }	        
//	if(document.frm1.txtespecM.value == "")
//      {
//      alert("É obrigatória o preenchimento do campo Especificação.");
//      document.frm1.txtespecM.focus();
//      return;
//      }	        
   else
      {
	   document.frm1.submit();
      }		
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="valida_alterar_micro.asp" name="frm1">
        <input type="hidden" name="txtOPT2" value="<%=str_OPT%>"><input type="hidden" name="txtOPT" value="<%=str_OPT%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
            <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28"></td>
            <td width="26">&nbsp;</td>
            <td width="159"></td>
          </tr>
        </table>
      </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="33">
    <tr> 
    <td> 
        <p align="center"><font face="Verdana" color="#330099" size="3">ALTERAÇÃO
        DE MICRO PERFIL</font>
    </td>
  </tr>
  <tr> 
      <td> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%'=str_Titulo%> </font></div>
      </td>
  </tr>
</table>
  <table width="968" border="0" cellspacing="0" cellpadding="0" height="45">
    <tr> 
      <td width="8" height="7">&nbsp;</td>
      <td width="217" height="7">&nbsp;</td>
      <td width="12" height="7"></td>
      <td width="724" height="7">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <%
	  'str_SQL = "SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO")
	  'response.Write(str_SQL)
	  'response.write rs("MEPR_CD_MEGA_PROCESSO")
	  
      SET TEMP=DB.EXECUTE("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
      %>
      <td width="724" height="35"><font face="Verdana" color="#330099" size="2"><%=UCASE(TEMP("MEPR_TX_DESC_MEGA_PROCESSO"))%></font></td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          R/3 :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <%
      SET TEMP=DB.EXECUTE("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO=	'" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
      %>
      <td width="724" height="35"><font face="Verdana" color="#330099" size="2"><%=UCASE(TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO"))%> </font></td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>ou 
          Macro Perfil :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <td width="724" height="35">
      <%
      SET TEMP=DB.EXECUTE("SELECT * FROM MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
      %>
        <p><font face="Verdana" color="#330099" size="2"><%=UCASE(TEMP("MCPE_TX_NOME_TECNICO"))%></font></p>
      </td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o 
          : </b></font></div></td>
      <td width="12" height="35"> </td>
      <td width="724" height="35"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=UCASE(TEMP("MCPE_TX_DESC_MACRO_PERFIL"))%></font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><input type="hidden" name="txtDesc" size="20" value="<%=str_Desc_Macro%>"></font></td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o
        Detalhada: </b></font></td>
      <td width="12" height="35"> 
      </td>
      <td width="724" height="35"> 
      <font face="Verdana" color="#330099" size="2"><%=UCASE(TEMP("MCPE_TX_DESC_DETA_MACRO_PERFIL"))%></font> 
      <font face="Verdana" color="#330099" size="1"><input type="hidden" name="txtdetal" size="20" value="<%=UCASE(TEMP("MCPE_TX_DESC_DETA_MACRO_PERFIL"))%>"></font> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Especificação
        :&nbsp;</b></font></td>
      <td width="12" height="35"> 
      </td>
      <td width="724" height="35"> 
        <font face="Verdana" color="#330099" size="2"><%=UCASE(TEMP("MCPE_TX_ESPECIFICACAO"))%></font> 
        <font face="Verdana" color="#330099" size="1"><input type="hidden" name="txtespec" size="20" value="<%=UCASE(TEMP("MCPE_TX_ESPECIFICACAO"))%>"></font> 
      </td>
    </tr>
  </table>
        <p><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  &nbsp;<span style="background-color: #C0C0C0">&nbsp;&nbsp;
  Detalhes do Micro Perfil&nbsp;&nbsp;&nbsp;&nbsp;</span></font></b></p>
  <table width="968" border="0" cellspacing="0" cellpadding="0" height="31">
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descrição:&nbsp;</b></font> 
        <p align="right"><input type="hidden" name="selMicro" size="20" value="<%=request("selMicroPerfil")%>"></td>
      <td width="502" height="35"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
          <textarea rows="3" name="txtDescM" cols="61" onKeyUp="pega_tamanho()"><%=RS("MICR_TX_DESC_MICRO_PERFIL")%></textarea>
          <input type="hidden" name="txtDescMicroPerfil_Original" size="20" value="<%=RS("MICR_TX_DESC_MICRO_PERFIL")%>">
        </p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="1">Tamanho
        Atual : <input type="text" name="txttamanho" size="5" maxlength="2" value="0">&nbsp;&nbsp;
        (Max
        61 Carateres)&nbsp;&nbsp;</font> 
      </td>
      <td width="222" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 
        </p>
      </td>
    </tr>
    <tr> 
      <td width="8" height="19"></td>
      <td width="217" height="19" valign="top">
      </td>
      <td width="502" height="19"> 
      </td>
      <td width="222" height="19"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o
        Detalhada: </b></font></td>
      <td width="502" height="35"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
          <textarea rows="2" name="txtdetalM" cols="61"><%=RS("MICR_TX_DESC_DETA_MICRO_PERFIL")%></textarea>
          <input type="hidden" name="txtDescDetalhada_Original" size="20" value="<%=RS("MICR_TX_DESC_DETA_MICRO_PERFIL")%>">
      </td>
      <td width="222" height="35"> 
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; 
      </td>
    </tr>
    <tr> 
      <td width="8" height="9"></td>
      <td width="217" height="9" valign="top">
      </td>
      <td width="502" height="9"> 
      </td>
      <td width="222" height="9"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="9"></td>
      <td width="217" height="9" valign="top">
      <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Histórico
      :&nbsp;</b></font>
      </td>
      <td width="502" height="9"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=RS("MICR_TX_ESPECIFICACAO")%></font> 
      </td>
      <td width="222" height="9"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="9"></td>
      <td width="217" height="9" valign="top">
      </td>
      <td width="502" height="9"> 
      <input type="hidden" name="Espec_ant" size="71" value="<%=RS("MICR_TX_ESPECIFICACAO")%>">
        <input type="hidden" name="txtEspecificacao_Original" size="20" value="<%=RS("MICR_TX_ESPECIFICACAO")%>"> 
      </td>
      <td width="222" height="9"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Especificação
        : </b></font></td>
      <td width="502" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><textarea rows="2" name="txtespecM" cols="61"></textarea> 
      </td>
      <td width="222" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; 
      </td>
    </tr>
  </table>
</form>
</body>
<script>
pega_tamanho()
</script>
</html>

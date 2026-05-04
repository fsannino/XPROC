<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL WHERE (MICR_TX_SITUACAO='EC' OR MICR_TX_SITUACAO = 'EL' OR MICR_TX_SITUACAO = 'AR' OR MICR_TX_SITUACAO = 'ER') ORDER BY MICR_TX_SEQ_MICRO_PERFIL")
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

function pega_caminho()
{
	var a = document.URL;
	var n=0;

	for (var i = 1 ; i < 1000; i++)
	{
	var final=a.slice(0,i)
	var t=a.slice(i-1,i);
	if (t=='/')
	{
	n = n + 1;
	}
	if(n == 4)
	{
	i = 1000;
	}
	}
	var tam=final.length;
	var caminho = final.slice(0,tam-1);
	
	document.frm1.txtcaminho.value=caminho;
	
}

function ver_historico(macro)
{
var a=macro;
window.open("ver_historico.asp?micro=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0")
}

</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="pega_caminho()" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="post" action="valida_micro2.asp">
        <font color="#330099" face="Verdana" size="3"><input type="hidden" name="usuario" size="10" value="<%=Session("CdUsuario")%>"></font>
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
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../Cenario/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
            <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Criação -&gt; Criado no R/3</font></p>
        
  <table border="0" width="868">
    <tr> 
      <td width="55" bgcolor="#330099" align="center" valign="middle"><b><font face="Verdana" size="1" color="#FFFFFF">Criado</font></b></td>
      <td width="71" bgcolor="#330099" align="center" valign="middle"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Recusado</font></b></div></td>
      <td width="74" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
          Altera&ccedil;&atilde;o R/3</font></b></div></td>
      <td width="74" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
          Exclus&atilde;o R/3</font></b></div></td>
      <td width="68" bgcolor="#330099" align="center" valign="middle"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Alterado 
          R/3</font></b></div></td>
      <td width="62" bgcolor="#330099" align="center" valign="middle"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Exclu&iacute;do 
          R/3</font></b></div></td>
      <td width="107" bgcolor="#330099" align="center" valign="middle"><b><font face="Verdana" size="1" color="#FFFFFF">Micro 
        - Perfil</font></b></td>
      <td width="194" bgcolor="#330099" align="center" valign="middle"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição</font></b></td>
      <td width="62" bgcolor="#FFFFFF"><font color="#330099" face="Verdana" size="2">&nbsp; 
        </font></td>
    </tr>
    <%
          VALOR_2=""
		   VALOR_3=""
		  
		  tem=0
          DO UNTIL RS.EOF=TRUE
		  
		  select case rs("MICR_TX_SITUACAO")
		  case "EL"
			valor_3="X"
		  case "RE"
			VALOR_2="checked"
		  	case "AR"
				VALOR_10="X"
			case "ER"
				VALOR_11="X"
			case "AP"
				VALOR_12="X"
			case "EP"
				VALOR_13="X"
		  end select
		  %>
    <tr> 
      <td width="55" height="24" align="center"> <p> 
          <%IF VALOR_3="" THEN%>
          <input type="radio" name="micro_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" value="1">
          <%END IF%>
          <font color="#330099" face="Verdana" size="1"> </font></p></td>
      <td width="71"><div align="center"> 
          <p> 
            <%IF VALOR_3="" THEN%>
            <input type="radio" name="micro_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" value="2" <%=VALOR_2%>>
            <%END IF%>
            <font color="#330099" face="Verdana" size="1"> </font></p>
        </div></td>
      <td width="74"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=VALOR_10%></strong></font></div></td>
      <td width="74"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=VALOR_11%></strong></font></div></td>
      <%
	  if valor_10="X" then
	  %>
	  <td width="68"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong><input type="radio" name="micro_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" value="3"></strong></font></div></td>
      <td width="62"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
	  <%else
	  if valor_11="X" then
	  %>
	  <td width="68"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
      <td width="62"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong><input type="radio" name="micro_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" value="4"></strong></font></div></td>
	  <%else%>
	  <td width="68"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
      <td width="62"><div align="center"><font color="#0000CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
	  <%
	  end if
	  end if
	  %>
      <td width="107"><font color="#330099" face="Verdana" size="1"><a href="gera_rel_micro_envia.asp?selMicro=<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>"><b><%=RS("MICR_TX_SEQ_MICRO_PERFIL")%></a></b></font></td>
      <td width="194"><font color="#330099" face="Verdana" size="1"><%=RS("MICR_TX_DESC_MICRO_PERFIL")%></font></td>
      <td width="62" bgcolor="#FFFFFF"> <p align="center"> <a href="#" onclick="ver_historico('<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>')"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></a> 
      </td>
    </tr>
    <tr> 
      <td height="53" colspan="2" align="center" valign="middle"><font color="#000033" size="2"><strong>Coment&aacute;rios/Motivo 
        :</strong></font></td>
      <td colspan="7" align="center" valign="middle"><div align="left"> 
          <textarea name="coment_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" cols="80" rows="2"></textarea>
          <font color="#330099" face="Verdana" size="2">
          <input type="hidden" name="mega_<%=trim(RS("MICR_TX_SEQ_MICRO_PERFIL"))%>" size="20" value=<%=rs("MEPR_CD_MEGA_PROCESSO")%>>
          </font> </div></td>
    </tr>
    <%
            tem=tem+1
            RS.MOVENEXT
			
			VALOR_2=""
			VALOR_3=""
			VALOR_10=""
			VALOR_11=""
			VALOR_12=""
			VALOR_13=""

            LOOP
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
